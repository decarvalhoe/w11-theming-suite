// TaskbarTAP.cpp — Taskbar Appearance Plugin for w11-theming-suite
// Two-stage injection architecture (same as TranslucentTB):
//   Stage 1: PowerShell injects this DLL into explorer.exe via CreateRemoteThread+LoadLibrary
//   Stage 2: DllMain spawns a thread that calls InitializeXamlDiagnosticsEx from WITHIN
//            explorer.exe, which triggers XAML Diagnostics to CoCreate our TAPSite.
//            TAPSite::SetSite receives IVisualTreeService3 and starts the VisualTreeWatcher.
//
// (c) 2026 w11-theming-suite. MIT License.

#include <initguid.h>   // Must come before guids.h to define (not just declare) GUIDs
#include "TaskbarTAP.h"
#include "guids.h"
#include <string>
#include <cstring>
#include <oleauto.h>   // SysAllocString, SysFreeString
#include <shlwapi.h>   // PathRemoveFileSpec (for getting DLL path)
#include <cstdio>      // for debug logging

#pragma comment(lib, "shlwapi.lib")

// ── Debug logging ──
static FILE* g_logFile = nullptr;

static void DebugLog(const char* fmt, ...)
{
    if (!g_logFile) {
        g_logFile = fopen("C:\\Dev\\w11-theming-suite\\native\\bin\\TaskbarTAP.log", "a");
        if (!g_logFile) return;
    }
    va_list args;
    va_start(args, fmt);
    vfprintf(g_logFile, fmt, args);
    va_end(args);
    fprintf(g_logFile, "\n");
    fflush(g_logFile);
}

// ── Globals ──
HMODULE g_hModule = nullptr;
std::atomic<long> g_refCount(0);
IVisualTreeService3* g_pTreeService = nullptr;
IXamlDiagnostics* g_pDiagnostics = nullptr;
// Default to TRANSPARENT — the whole point of injecting this DLL
TaskbarAppearance g_appearance = APPEARANCE_TRANSPARENT;
VisualTreeWatcher* g_pWatcher = nullptr;

// ── Shared memory for IPC (PowerShell writes, DLL reads) ──
static const wchar_t* SHARED_MEM_NAME = L"W11ThemeSuite_TaskbarTAP_Mode";
static HANDLE g_hMapFile = nullptr;
static volatile int* g_pSharedMode = nullptr;
static HANDLE g_hMonitorThread = nullptr;
static bool g_bStopMonitor = false;

static void InitSharedMemory()
{
    g_hMapFile = CreateFileMappingW(
        INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0, sizeof(int), SHARED_MEM_NAME);
    if (g_hMapFile) {
        g_pSharedMode = (volatile int*)MapViewOfFile(
            g_hMapFile, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(int));
        if (g_pSharedMode) {
            *g_pSharedMode = (int)g_appearance;
        }
    }
}

// Monitor thread: polls shared memory for mode changes from PowerShell
static DWORD WINAPI MonitorThread(LPVOID)
{
    while (!g_bStopMonitor) {
        if (g_pSharedMode) {
            int newMode = *g_pSharedMode;
            if (newMode >= 0 && newMode <= 2 && newMode != (int)g_appearance) {
                g_appearance = (TaskbarAppearance)newMode;
                if (g_pWatcher) {
                    g_pWatcher->ApplyAppearance(g_appearance);
                }
            }
        }
        Sleep(250);  // Check 4 times per second
    }
    return 0;
}

static void StartMonitorThread()
{
    g_bStopMonitor = false;
    g_hMonitorThread = CreateThread(nullptr, 0, MonitorThread, nullptr, 0, nullptr);
}

// ══════════════════════════════════════════════
// Stage 2: Self-injection into XAML Diagnostics
// This runs inside explorer.exe after LoadLibrary injection.
// Dynamically loads Windows.UI.Xaml.dll, resolves InitializeXamlDiagnosticsEx,
// and calls it with our own DLL path and CLSID.
// ══════════════════════════════════════════════
typedef HRESULT(WINAPI* PFN_InitializeXamlDiagnosticsEx)(
    LPCWSTR endPointName,
    DWORD pid,
    LPCWSTR wszDllXamlDiagnostics,
    LPCWSTR wszTAPDllName,
    CLSID tapClsid,
    LPCWSTR wszInitializationData
);

static DWORD WINAPI SelfInjectThread(LPVOID)
{
    DebugLog("=== SelfInjectThread started ===");

    // Get our own DLL path
    wchar_t dllPath[MAX_PATH];
    GetModuleFileNameW(g_hModule, dllPath, MAX_PATH);
    DebugLog("DLL path: %ls", dllPath);

    // Load Windows.UI.Xaml.dll and resolve InitializeXamlDiagnosticsEx
    HMODULE hWux = LoadLibraryExW(L"Windows.UI.Xaml.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!hWux) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    auto pfnIXDE = reinterpret_cast<PFN_InitializeXamlDiagnosticsEx>(
        GetProcAddress(hWux, "InitializeXamlDiagnosticsEx"));
    if (!pfnIXDE) {
        FreeLibrary(hWux);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD pid = GetCurrentProcessId();
    HRESULT hr = E_FAIL;
    UINT8 attempts = 1;

    // Retry loop: each attempt uses a new thread because IXDE is once-per-thread.
    // Use "VisualDiagConnection{N}" endpoint names (same as TranslucentTB).
    do {
        wchar_t connName[64];
        wsprintfW(connName, L"VisualDiagConnection%d", (int)attempts);

        // Must call IXDE on a fresh thread (it's once-per-thread)
        struct IxdeArgs {
            PFN_InitializeXamlDiagnosticsEx pfn;
            LPCWSTR conn;
            DWORD pid;
            LPCWSTR dllPath;
            HRESULT hr;
        };

        IxdeArgs args;
        args.pfn = pfnIXDE;
        args.conn = connName;
        args.pid = pid;
        args.dllPath = dllPath;
        args.hr = E_FAIL;

        HANDLE hThread = CreateThread(nullptr, 0, [](LPVOID param) -> DWORD {
            auto* a = (IxdeArgs*)param;
            a->hr = a->pfn(a->conn, a->pid, nullptr, a->dllPath, CLSID_TaskbarTAPSite, nullptr);
            return 0;
        }, &args, 0, nullptr);

        if (hThread) {
            WaitForSingleObject(hThread, 5000);
            CloseHandle(hThread);
            hr = args.hr;
        }

        if (SUCCEEDED(hr)) {
            DebugLog("IXDE succeeded on attempt %d", (int)attempts);
            break;
        }

        DebugLog("IXDE attempt %d failed: 0x%08X", (int)attempts, hr);
        ++attempts;
        Sleep(500);
    } while (FAILED(hr) && attempts <= 60);

    if (FAILED(hr)) {
        DebugLog("IXDE FAILED after all attempts. Last HRESULT: 0x%08X", hr);
    }

    return hr;
}

// ══════════════════════════════════════════════
// DLL Entry Point
// ══════════════════════════════════════════════
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH) {
        g_hModule = hInstance;
        DisableThreadLibraryCalls(hInstance);

        // Stage 2: Spawn self-injection thread.
        // This will call InitializeXamlDiagnosticsEx from WITHIN explorer.exe.
        HANDLE hThread = CreateThread(nullptr, 0, SelfInjectThread, nullptr, 0, nullptr);
        if (hThread) {
            CloseHandle(hThread);
        }
    }
    else if (reason == DLL_PROCESS_DETACH) {
        g_bStopMonitor = true;
        if (g_hMonitorThread) {
            WaitForSingleObject(g_hMonitorThread, 2000);
            CloseHandle(g_hMonitorThread);
            g_hMonitorThread = nullptr;
        }
        if (g_pSharedMode) { UnmapViewOfFile((LPCVOID)g_pSharedMode); g_pSharedMode = nullptr; }
        if (g_hMapFile) { CloseHandle(g_hMapFile); g_hMapFile = nullptr; }
    }
    return TRUE;
}

// ══════════════════════════════════════════════
// Exported functions (for IPC/version check)
// ══════════════════════════════════════════════
extern "C" {

HRESULT __stdcall SetTaskbarTransparent()
{
    g_appearance = APPEARANCE_TRANSPARENT;
    if (g_pWatcher) g_pWatcher->ApplyAppearance(g_appearance);
    return S_OK;
}

HRESULT __stdcall SetTaskbarAcrylic()
{
    g_appearance = APPEARANCE_ACRYLIC;
    if (g_pWatcher) g_pWatcher->ApplyAppearance(g_appearance);
    return S_OK;
}

HRESULT __stdcall SetTaskbarDefault()
{
    g_appearance = APPEARANCE_DEFAULT;
    if (g_pWatcher) g_pWatcher->ApplyAppearance(g_appearance);
    return S_OK;
}

int __stdcall GetTaskbarTAPVersion()
{
    return 1;  // v1.0
}

} // extern "C"

// ══════════════════════════════════════════════
// COM: DllGetClassObject + DllCanUnloadNow
// ══════════════════════════════════════════════
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    if (rclsid == CLSID_TaskbarTAPSite) {
        auto* factory = new TaskbarTAPFactory();
        HRESULT hr = factory->QueryInterface(riid, ppv);
        factory->Release();
        return hr;
    }
    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow()
{
    return (g_refCount == 0) ? S_OK : S_FALSE;
}

// ══════════════════════════════════════════════
// TaskbarTAPFactory (IClassFactory)
// ══════════════════════════════════════════════
TaskbarTAPFactory::TaskbarTAPFactory() : m_refCount(1) { g_refCount++; }

HRESULT TaskbarTAPFactory::QueryInterface(REFIID riid, void** ppv)
{
    if (riid == IID_IUnknown || riid == IID_IClassFactory) {
        *ppv = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
}

ULONG TaskbarTAPFactory::AddRef() { return InterlockedIncrement(&m_refCount); }
ULONG TaskbarTAPFactory::Release()
{
    long ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) { g_refCount--; delete this; }
    return ref;
}

HRESULT TaskbarTAPFactory::CreateInstance(IUnknown* pOuter, REFIID riid, void** ppv)
{
    if (pOuter) return CLASS_E_NOAGGREGATION;
    auto* site = new TaskbarTAPSite();
    HRESULT hr = site->QueryInterface(riid, ppv);
    site->Release();
    return hr;
}

HRESULT TaskbarTAPFactory::LockServer(BOOL fLock)
{
    if (fLock) g_refCount++;
    else g_refCount--;
    return S_OK;
}

// ══════════════════════════════════════════════
// TaskbarTAPSite (IObjectWithSite)
// XAML Diagnostics calls SetSite() with the diagnostics provider.
// We QI for IVisualTreeService3 and IXamlDiagnostics, then start
// watching the visual tree.
// ══════════════════════════════════════════════
TaskbarTAPSite::TaskbarTAPSite() : m_refCount(1), m_pSite(nullptr)
{
    g_refCount++;
}

TaskbarTAPSite::~TaskbarTAPSite()
{
    if (m_pSite) m_pSite->Release();
    g_refCount--;
}

HRESULT TaskbarTAPSite::QueryInterface(REFIID riid, void** ppv)
{
    if (riid == IID_IUnknown || riid == IID_IObjectWithSite) {
        *ppv = static_cast<IObjectWithSite*>(this);
        AddRef();
        return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
}

ULONG TaskbarTAPSite::AddRef() { return InterlockedIncrement(&m_refCount); }
ULONG TaskbarTAPSite::Release()
{
    long ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) delete this;
    return ref;
}

HRESULT TaskbarTAPSite::SetSite(IUnknown* pUnkSite)
{
    DebugLog("=== SetSite called (pUnkSite=%p) ===", pUnkSite);

    // Release previous site
    if (m_pSite) { m_pSite->Release(); m_pSite = nullptr; }
    if (g_pTreeService) { g_pTreeService->Release(); g_pTreeService = nullptr; }
    if (g_pDiagnostics) { g_pDiagnostics->Release(); g_pDiagnostics = nullptr; }
    if (g_pWatcher) { g_pWatcher->Release(); g_pWatcher = nullptr; }

    if (!pUnkSite) return S_OK;  // Disconnecting

    m_pSite = pUnkSite;
    m_pSite->AddRef();

    // QI for the XAML diagnostics interfaces
    HRESULT hr = pUnkSite->QueryInterface(__uuidof(IXamlDiagnostics),
                                           reinterpret_cast<void**>(&g_pDiagnostics));
    DebugLog("QI IXamlDiagnostics: 0x%08X (ptr=%p)", hr, g_pDiagnostics);
    if (FAILED(hr)) return hr;

    hr = pUnkSite->QueryInterface(__uuidof(IVisualTreeService3),
                                   reinterpret_cast<void**>(&g_pTreeService));
    DebugLog("QI IVisualTreeService3: 0x%08X (ptr=%p)", hr, g_pTreeService);
    if (FAILED(hr)) return hr;

    // Create watcher and subscribe to visual tree changes
    // Run AdviseVisualTreeChange on a new thread to avoid blocking
    g_pWatcher = new VisualTreeWatcher(g_pDiagnostics, g_pTreeService);

    HANDLE hThread = CreateThread(nullptr, 0, [](LPVOID param) -> DWORD {
        auto* watcher = (VisualTreeWatcher*)param;
        HRESULT hr = g_pTreeService->AdviseVisualTreeChange(watcher);
        return hr;
    }, g_pWatcher, 0, nullptr);

    if (hThread) {
        CloseHandle(hThread);
    }

    // Initialize shared memory for IPC with PowerShell
    InitSharedMemory();
    StartMonitorThread();

    return S_OK;
}

HRESULT TaskbarTAPSite::GetSite(REFIID riid, void** ppvSite)
{
    if (!m_pSite) {
        *ppvSite = nullptr;
        return E_FAIL;
    }
    return m_pSite->QueryInterface(riid, ppvSite);
}

// ══════════════════════════════════════════════
// VisualTreeWatcher (IVisualTreeServiceCallback2)
// Watches XAML element creation/mutation to find and modify
// the taskbar background elements.
// ══════════════════════════════════════════════
VisualTreeWatcher::VisualTreeWatcher(IXamlDiagnostics* pDiag, IVisualTreeService3* pService)
    : m_refCount(1), m_pDiag(pDiag), m_pService(pService), m_taskbarCount(0)
{
    g_refCount++;
    if (m_pDiag) m_pDiag->AddRef();
    if (m_pService) m_pService->AddRef();
    memset(m_taskbars, 0, sizeof(m_taskbars));
}

VisualTreeWatcher::~VisualTreeWatcher()
{
    if (m_pDiag) m_pDiag->Release();
    if (m_pService) m_pService->Release();
    g_refCount--;
}

HRESULT VisualTreeWatcher::QueryInterface(REFIID riid, void** ppv)
{
    if (riid == IID_IUnknown ||
        riid == __uuidof(IVisualTreeServiceCallback) ||
        riid == __uuidof(IVisualTreeServiceCallback2)) {
        *ppv = static_cast<IVisualTreeServiceCallback2*>(this);
        AddRef();
        return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
}

ULONG VisualTreeWatcher::AddRef() { return InterlockedIncrement(&m_refCount); }
ULONG VisualTreeWatcher::Release()
{
    long ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) delete this;
    return ref;
}

// ── OnVisualTreeChange ──
// Called by XAML Diagnostics whenever an element is added/removed/changed.
// We look for Rectangle#BackgroundFill and Rectangle#BackgroundStroke.
HRESULT VisualTreeWatcher::OnVisualTreeChange(
    ParentChildRelation relation,
    VisualElement element,
    VisualMutationType mutationType)
{
    if (mutationType == Add) {
        if (element.Name && element.Type) {
            std::wstring name(element.Name);
            std::wstring type(element.Type);

            bool isBackgroundFill = (name == L"BackgroundFill" &&
                                      type.find(L"Rectangle") != std::wstring::npos);
            bool isBackgroundStroke = (name == L"BackgroundStroke" &&
                                       type.find(L"Rectangle") != std::wstring::npos);
            bool isTaskbarFrame = (type.find(L"TaskbarFrame") != std::wstring::npos);

            if (isBackgroundFill || isBackgroundStroke || isTaskbarFrame) {
                DebugLog("Found element: name='%ls' type='%ls' handle=%llu",
                    name.c_str(), type.c_str(), (unsigned long long)element.Handle);
                int slot = -1;

                if (isBackgroundFill || isBackgroundStroke) {
                    for (int i = 0; i < m_taskbarCount; i++) {
                        if (m_taskbars[i].active) {
                            if (isBackgroundFill && m_taskbars[i].backgroundFill == 0) {
                                slot = i; break;
                            }
                            if (isBackgroundStroke && m_taskbars[i].backgroundStroke == 0) {
                                slot = i; break;
                            }
                        }
                    }
                }

                if (slot < 0 && m_taskbarCount < MAX_TASKBARS) {
                    slot = m_taskbarCount++;
                    m_taskbars[slot].active = true;
                    m_taskbars[slot].backgroundFill = 0;
                    m_taskbars[slot].backgroundStroke = 0;
                    m_taskbars[slot].taskbarFrame = 0;
                }

                if (slot >= 0) {
                    if (isBackgroundFill) {
                        m_taskbars[slot].backgroundFill = element.Handle;
                    } else if (isBackgroundStroke) {
                        m_taskbars[slot].backgroundStroke = element.Handle;
                    } else if (isTaskbarFrame) {
                        m_taskbars[slot].taskbarFrame = element.Handle;
                    }

                    // Apply current appearance immediately
                    if (g_appearance != APPEARANCE_DEFAULT) {
                        ApplyAppearance(g_appearance);
                    }
                }
            }
        }
    }
    else if (mutationType == Remove) {
        for (int i = 0; i < m_taskbarCount; i++) {
            if (m_taskbars[i].backgroundFill == element.Handle)
                m_taskbars[i].backgroundFill = 0;
            if (m_taskbars[i].backgroundStroke == element.Handle)
                m_taskbars[i].backgroundStroke = 0;
            if (m_taskbars[i].taskbarFrame == element.Handle)
                m_taskbars[i].taskbarFrame = 0;

            if (m_taskbars[i].active &&
                m_taskbars[i].backgroundFill == 0 &&
                m_taskbars[i].backgroundStroke == 0 &&
                m_taskbars[i].taskbarFrame == 0) {
                m_taskbars[i].active = false;
            }
        }
    }

    return S_OK;
}

HRESULT VisualTreeWatcher::OnElementStateChanged(
    InstanceHandle /*element*/,
    VisualElementState /*elementState*/,
    LPCWSTR /*context*/)
{
    return S_OK;
}

// ── Helper: Set a XAML property via the correct 3-step API ──
static HRESULT SetXamlProperty(IVisualTreeService3* pService,
                                InstanceHandle element,
                                LPCWSTR propertyName,
                                LPCWSTR typeName,
                                LPCWSTR value)
{
    if (!pService || element == 0) return E_INVALIDARG;

    // Step 1: Create the value instance
    InstanceHandle hValue = 0;
    BSTR bstrType = SysAllocString(typeName);
    BSTR bstrVal  = value ? SysAllocString(value) : nullptr;
    HRESULT hr = pService->CreateInstance(bstrType, bstrVal, &hValue);
    SysFreeString(bstrType);
    if (bstrVal) SysFreeString(bstrVal);
    DebugLog("  SetXamlProperty: CreateInstance('%ls','%ls') = 0x%08X (handle=%llu)",
        typeName, value ? value : L"(null)", hr, (unsigned long long)hValue);
    if (FAILED(hr)) return hr;

    // Step 2: Get property index
    unsigned int propIndex = 0;
    hr = pService->GetPropertyIndex(element, propertyName, &propIndex);
    DebugLog("  SetXamlProperty: GetPropertyIndex('%ls') = 0x%08X (index=%u)",
        propertyName, hr, propIndex);
    if (FAILED(hr)) return hr;

    // Step 3: Set the property
    hr = pService->SetProperty(element, hValue, propIndex);
    DebugLog("  SetXamlProperty: SetProperty(elem=%llu, val=%llu, idx=%u) = 0x%08X",
        (unsigned long long)element, (unsigned long long)hValue, propIndex, hr);
    return hr;
}

// Helper: Clear a property (revert to default binding)
static HRESULT ClearXamlProperty(IVisualTreeService3* pService,
                                  InstanceHandle element,
                                  LPCWSTR propertyName)
{
    if (!pService || element == 0) return E_INVALIDARG;

    unsigned int propIndex = 0;
    HRESULT hr = pService->GetPropertyIndex(element, propertyName, &propIndex);
    if (FAILED(hr)) return hr;

    return pService->ClearProperty(element, propIndex);
}

// ── ApplyAppearance ──
void VisualTreeWatcher::ApplyAppearance(TaskbarAppearance appearance)
{
    DebugLog("ApplyAppearance called: mode=%d taskbarCount=%d", (int)appearance, m_taskbarCount);
    if (!m_pService) { DebugLog("  ERROR: m_pService is null!"); return; }

    for (int i = 0; i < m_taskbarCount; i++) {
        if (!m_taskbars[i].active) continue;

        InstanceHandle bgFill = m_taskbars[i].backgroundFill;
        InstanceHandle bgStroke = m_taskbars[i].backgroundStroke;

        if (bgFill != 0) {
            switch (appearance) {
            case APPEARANCE_TRANSPARENT:
                SetXamlProperty(m_pService, bgFill, L"Fill",
                    L"Windows.UI.Xaml.Media.SolidColorBrush", L"Transparent");
                SetXamlProperty(m_pService, bgFill, L"Opacity",
                    L"Double", L"0");
                break;

            case APPEARANCE_ACRYLIC:
                SetXamlProperty(m_pService, bgFill, L"Fill",
                    L"Windows.UI.Xaml.Media.SolidColorBrush", L"#44000000");
                SetXamlProperty(m_pService, bgFill, L"Opacity",
                    L"Double", L"1");
                break;

            case APPEARANCE_DEFAULT:
                ClearXamlProperty(m_pService, bgFill, L"Fill");
                ClearXamlProperty(m_pService, bgFill, L"Opacity");
                break;
            }
        }

        if (bgStroke != 0) {
            switch (appearance) {
            case APPEARANCE_TRANSPARENT:
                SetXamlProperty(m_pService, bgStroke, L"Fill",
                    L"Windows.UI.Xaml.Media.SolidColorBrush", L"Transparent");
                SetXamlProperty(m_pService, bgStroke, L"Opacity",
                    L"Double", L"0");
                break;

            case APPEARANCE_ACRYLIC:
                SetXamlProperty(m_pService, bgStroke, L"Fill",
                    L"Windows.UI.Xaml.Media.SolidColorBrush", L"Transparent");
                SetXamlProperty(m_pService, bgStroke, L"Opacity",
                    L"Double", L"0");
                break;

            case APPEARANCE_DEFAULT:
                ClearXamlProperty(m_pService, bgStroke, L"Fill");
                ClearXamlProperty(m_pService, bgStroke, L"Opacity");
                break;
            }
        }
    }
}
