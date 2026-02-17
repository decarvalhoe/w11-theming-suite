// ShellTAP.cpp -- Generic Shell Transparency/Appearance Plugin for w11-theming-suite
// Refactored from TaskbarTAP.cpp to support injection into any XAML process.
//
// Two-stage injection architecture (same as TranslucentTB):
//   Stage 1: PowerShell injects this DLL via CreateRemoteThread + LoadLibraryW
//   Stage 2: DllMain spawns a thread calling InitializeXamlDiagnosticsEx from
//            within the target process. XAML Diagnostics then CoCreates our
//            ShellTAPSite, which starts the VisualTreeWatcher.
//
// Configuration is read from named shared memory:
//   "W11ThemeSuite_ShellTAP_<TargetId>_Config" -- ShellTAPConfig struct
//   "W11ThemeSuite_ShellTAP_<TargetId>_Mode"   -- int (mode changes from PS)
//
// If no config shared memory exists, operates in discovery mode (logs all elements).
//
// (c) 2026 w11-theming-suite. MIT License.

#include <initguid.h>   // Must come before guids.h
#include "ShellTAP.h"
#include "guids.h"
#include <string>
#include <cstring>
#include <oleauto.h>     // SysAllocString, SysFreeString
#include <shlwapi.h>     // PathRemoveFileSpec
#include <cstdio>        // debug logging
#include <winstring.h>   // WindowsGetStringRawBuffer, WindowsDeleteString

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "WindowsApp.lib")

// ── Debug logging ──
static FILE* g_logFile = nullptr;
static wchar_t g_logPath[MAX_PATH] = {0};

static void DebugLog(const char* fmt, ...)
{
    if (!g_logFile) {
        if (g_logPath[0] != 0) {
            g_logFile = _wfopen(g_logPath, L"a");
        } else {
            // Fallback: log next to the DLL
            wchar_t dllDir[MAX_PATH];
            GetModuleFileNameW(g_hModule, dllDir, MAX_PATH);
            PathRemoveFileSpecW(dllDir);
            wchar_t logPath[MAX_PATH];
            wsprintfW(logPath, L"%s\\ShellTAP.log", dllDir);
            g_logFile = _wfopen(logPath, L"a");
        }
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
AppearanceMode g_mode = MODE_TRANSPARENT;
VisualTreeWatcher* g_pWatcher = nullptr;

// ── Configuration from shared memory ──
static ShellTAPConfig g_config = {0};
static bool g_discoveryMode = true;  // default: discovery mode
static wchar_t g_targetId[64] = L"Unknown";

// ── Shared memory for mode IPC ──
static HANDLE g_hModeMap = nullptr;
static volatile int* g_pSharedMode = nullptr;
static HANDLE g_hMonitorThread = nullptr;
static bool g_bStopMonitor = false;

// ── Discovery mode log ──
static FILE* g_discoveryLog = nullptr;

static void DiscoveryLog(const char* fmt, ...)
{
    if (!g_discoveryLog) return;
    va_list args;
    va_start(args, fmt);
    vfprintf(g_discoveryLog, fmt, args);
    va_end(args);
    fprintf(g_discoveryLog, "\n");
    fflush(g_discoveryLog);
}

// Read configuration from shared memory (written by PowerShell before injection)
static bool ReadConfig()
{
    wchar_t configName[128];
    wsprintfW(configName, L"W11ThemeSuite_ShellTAP_%s_Config", g_targetId);

    HANDLE hMap = OpenFileMappingW(FILE_MAP_READ, FALSE, configName);
    if (!hMap) {
        DebugLog("No config shared memory '%ls' -- entering discovery mode", configName);
        return false;
    }

    void* pView = MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, sizeof(ShellTAPConfig));
    if (pView) {
        memcpy(&g_config, pView, sizeof(ShellTAPConfig));
        UnmapViewOfFile(pView);
    }
    CloseHandle(hMap);

    if (g_config.version != SHELLTAP_CONFIG_VERSION) {
        DebugLog("Config version mismatch: expected %d, got %d", SHELLTAP_CONFIG_VERSION, g_config.version);
        return false;
    }

    g_mode = (AppearanceMode)g_config.mode;
    g_discoveryMode = (g_config.targetCount == 0);

    if (g_config.logPath[0] != 0) {
        wcscpy_s(g_logPath, g_config.logPath);
    }

    DebugLog("Config loaded: mode=%d, targetCount=%d, discovery=%s",
        g_config.mode, g_config.targetCount, g_discoveryMode ? "YES" : "NO");

    for (int i = 0; i < g_config.targetCount && i < 8; i++) {
        DebugLog("  Target[%d]: name='%ls' type='%ls'",
            i, g_config.targetNames[i], g_config.targetTypes[i]);
    }

    return true;
}

// Initialize mode IPC shared memory
static void InitModeSharedMemory()
{
    wchar_t modeName[128];
    wsprintfW(modeName, L"W11ThemeSuite_ShellTAP_%s_Mode", g_targetId);

    g_hModeMap = CreateFileMappingW(
        INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0, sizeof(int), modeName);
    if (g_hModeMap) {
        g_pSharedMode = (volatile int*)MapViewOfFile(
            g_hModeMap, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(int));
        if (g_pSharedMode) {
            *g_pSharedMode = (int)g_mode;
        }
    }
    DebugLog("Mode shared memory '%ls' initialized (ptr=%p)", modeName, g_pSharedMode);
}

// Monitor thread: polls shared memory for mode changes from PowerShell
static DWORD WINAPI MonitorThread(LPVOID)
{
    while (!g_bStopMonitor) {
        if (g_pSharedMode) {
            int newMode = *g_pSharedMode;
            if (newMode >= 0 && newMode <= 2 && newMode != (int)g_mode) {
                g_mode = (AppearanceMode)newMode;
                DebugLog("Mode changed to %d via shared memory", newMode);
                if (g_pWatcher) {
                    g_pWatcher->ApplyMode(g_mode);
                }
            }
        }
        Sleep(250);
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
    DebugLog("=== SelfInjectThread started (target=%ls) ===", g_targetId);

    wchar_t dllPath[MAX_PATH];
    GetModuleFileNameW(g_hModule, dllPath, MAX_PATH);
    DebugLog("DLL path: %ls", dllPath);

    HMODULE hWux = LoadLibraryExW(L"Windows.UI.Xaml.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!hWux) {
        DebugLog("LoadLibrary(Windows.UI.Xaml.dll) FAILED: 0x%08X", GetLastError());
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

    do {
        wchar_t connName[64];
        wsprintfW(connName, L"VisualDiagConnection%d", (int)attempts);

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
            a->hr = a->pfn(a->conn, a->pid, nullptr, a->dllPath, CLSID_ShellTAPSite, nullptr);
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

        // Read TargetId from shared memory (written by PowerShell before injection)
        // Fixed name: "W11ThemeSuite_ShellTAP_Init" contains the target ID string
        HANDLE hInitMap = OpenFileMappingW(FILE_MAP_READ, FALSE, L"W11ThemeSuite_ShellTAP_Init");
        if (hInitMap) {
            void* pView = MapViewOfFile(hInitMap, FILE_MAP_READ, 0, 0, 64 * sizeof(wchar_t));
            if (pView) {
                wcsncpy_s(g_targetId, (const wchar_t*)pView, 63);
                UnmapViewOfFile(pView);
            }
            CloseHandle(hInitMap);
        }

        // Read configuration from shared memory
        ReadConfig();

        // Open discovery log if in discovery mode
        if (g_discoveryMode) {
            wchar_t discLogPath[MAX_PATH];
            if (g_config.logPath[0] != 0) {
                wsprintfW(discLogPath, L"%s", g_config.logPath);
            } else {
                wchar_t dllDir[MAX_PATH];
                GetModuleFileNameW(g_hModule, dllDir, MAX_PATH);
                PathRemoveFileSpecW(dllDir);
                wsprintfW(discLogPath, L"%s\\ShellTAP_%s_discovery.log", dllDir, g_targetId);
            }
            g_discoveryLog = _wfopen(discLogPath, L"w");
            if (g_discoveryLog) {
                fprintf(g_discoveryLog, "=== ShellTAP Discovery Log (target=%ls) ===\n", g_targetId);
                fprintf(g_discoveryLog, "Format: [handle] name | type\n\n");
                fflush(g_discoveryLog);
            }
        }

        // Spawn self-injection thread
        HANDLE hThread = CreateThread(nullptr, 0, SelfInjectThread, nullptr, 0, nullptr);
        if (hThread) CloseHandle(hThread);
    }
    else if (reason == DLL_PROCESS_DETACH) {
        g_bStopMonitor = true;
        if (g_hMonitorThread) {
            WaitForSingleObject(g_hMonitorThread, 2000);
            CloseHandle(g_hMonitorThread);
            g_hMonitorThread = nullptr;
        }
        if (g_pSharedMode) { UnmapViewOfFile((LPCVOID)g_pSharedMode); g_pSharedMode = nullptr; }
        if (g_hModeMap) { CloseHandle(g_hModeMap); g_hModeMap = nullptr; }
        if (g_discoveryLog) { fclose(g_discoveryLog); g_discoveryLog = nullptr; }
        if (g_logFile) { fclose(g_logFile); g_logFile = nullptr; }
    }
    return TRUE;
}

// ══════════════════════════════════════════════
// Exported functions
// ══════════════════════════════════════════════
extern "C" {

HRESULT __stdcall SetShellTAPMode(int mode)
{
    if (mode < 0 || mode > 2) return E_INVALIDARG;
    g_mode = (AppearanceMode)mode;
    if (g_pSharedMode) *g_pSharedMode = mode;
    if (g_pWatcher) g_pWatcher->ApplyMode(g_mode);
    return S_OK;
}

int __stdcall GetShellTAPMode()
{
    return (int)g_mode;
}

int __stdcall GetShellTAPVersion()
{
    return 1;
}

int __stdcall GetShellTAPAppliedCount()
{
    return g_pWatcher ? g_pWatcher->GetTrackedCount() : 0;
}

} // extern "C"

// ══════════════════════════════════════════════
// COM: DllGetClassObject + DllCanUnloadNow
// ══════════════════════════════════════════════
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    if (rclsid == CLSID_ShellTAPSite) {
        auto* factory = new ShellTAPFactory();
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
// ShellTAPFactory (IClassFactory)
// ══════════════════════════════════════════════
ShellTAPFactory::ShellTAPFactory() : m_refCount(1) { g_refCount++; }

HRESULT ShellTAPFactory::QueryInterface(REFIID riid, void** ppv)
{
    if (riid == IID_IUnknown || riid == IID_IClassFactory) {
        *ppv = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
}

ULONG ShellTAPFactory::AddRef() { return InterlockedIncrement(&m_refCount); }
ULONG ShellTAPFactory::Release()
{
    long ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) { g_refCount--; delete this; }
    return ref;
}

HRESULT ShellTAPFactory::CreateInstance(IUnknown* pOuter, REFIID riid, void** ppv)
{
    if (pOuter) return CLASS_E_NOAGGREGATION;
    auto* site = new ShellTAPSite();
    HRESULT hr = site->QueryInterface(riid, ppv);
    site->Release();
    return hr;
}

HRESULT ShellTAPFactory::LockServer(BOOL fLock)
{
    if (fLock) g_refCount++;
    else g_refCount--;
    return S_OK;
}

// ══════════════════════════════════════════════
// ShellTAPSite (IObjectWithSite)
// ══════════════════════════════════════════════
ShellTAPSite::ShellTAPSite() : m_refCount(1), m_pSite(nullptr) { g_refCount++; }

ShellTAPSite::~ShellTAPSite()
{
    if (m_pSite) m_pSite->Release();
    g_refCount--;
}

HRESULT ShellTAPSite::QueryInterface(REFIID riid, void** ppv)
{
    if (riid == IID_IUnknown || riid == IID_IObjectWithSite) {
        *ppv = static_cast<IObjectWithSite*>(this);
        AddRef();
        return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
}

ULONG ShellTAPSite::AddRef() { return InterlockedIncrement(&m_refCount); }
ULONG ShellTAPSite::Release()
{
    long ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) delete this;
    return ref;
}

HRESULT ShellTAPSite::SetSite(IUnknown* pUnkSite)
{
    DebugLog("=== SetSite called (target=%ls, pUnkSite=%p) ===", g_targetId, pUnkSite);

    if (m_pSite) { m_pSite->Release(); m_pSite = nullptr; }
    if (g_pTreeService) { g_pTreeService->Release(); g_pTreeService = nullptr; }
    if (g_pDiagnostics) { g_pDiagnostics->Release(); g_pDiagnostics = nullptr; }
    if (g_pWatcher) { g_pWatcher->Release(); g_pWatcher = nullptr; }

    if (!pUnkSite) return S_OK;

    m_pSite = pUnkSite;
    m_pSite->AddRef();

    HRESULT hr = pUnkSite->QueryInterface(__uuidof(IXamlDiagnostics),
                                           reinterpret_cast<void**>(&g_pDiagnostics));
    DebugLog("QI IXamlDiagnostics: 0x%08X", hr);
    if (FAILED(hr)) return hr;

    hr = pUnkSite->QueryInterface(__uuidof(IVisualTreeService3),
                                   reinterpret_cast<void**>(&g_pTreeService));
    DebugLog("QI IVisualTreeService3: 0x%08X", hr);
    if (FAILED(hr)) return hr;

    g_pWatcher = new VisualTreeWatcher(g_pDiagnostics, g_pTreeService);

    HANDLE hThread = CreateThread(nullptr, 0, [](LPVOID param) -> DWORD {
        auto* watcher = (VisualTreeWatcher*)param;
        HRESULT hr = g_pTreeService->AdviseVisualTreeChange(watcher);
        DebugLog("AdviseVisualTreeChange: 0x%08X", hr);
        return hr;
    }, g_pWatcher, 0, nullptr);

    if (hThread) CloseHandle(hThread);

    InitModeSharedMemory();
    StartMonitorThread();

    return S_OK;
}

HRESULT ShellTAPSite::GetSite(REFIID riid, void** ppvSite)
{
    if (!m_pSite) { *ppvSite = nullptr; return E_FAIL; }
    return m_pSite->QueryInterface(riid, ppvSite);
}

// ══════════════════════════════════════════════
// VisualTreeWatcher
// ══════════════════════════════════════════════
VisualTreeWatcher::VisualTreeWatcher(IXamlDiagnostics* pDiag, IVisualTreeService3* pService)
    : m_refCount(1), m_pDiag(pDiag), m_pService(pService), m_trackedCount(0)
{
    g_refCount++;
    if (m_pDiag) m_pDiag->AddRef();
    if (m_pService) m_pService->AddRef();
    memset(m_tracked, 0, sizeof(m_tracked));
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

// Check if an element matches any configured target
bool VisualTreeWatcher::MatchesTarget(const wchar_t* name, const wchar_t* type, bool* outIsStroke)
{
    *outIsStroke = false;

    for (int i = 0; i < g_config.targetCount && i < 8; i++) {
        bool nameMatch = false;
        bool typeMatch = false;

        // Name matching: exact match or wildcard "*"
        if (g_config.targetNames[i][0] == L'*' && g_config.targetNames[i][1] == 0) {
            nameMatch = true;
        } else if (name && wcscmp(name, g_config.targetNames[i]) == 0) {
            nameMatch = true;
        }

        // Type matching: substring match (e.g., "Rectangle" matches "Windows.UI.Xaml.Shapes.Rectangle")
        if (g_config.targetTypes[i][0] == L'*' && g_config.targetTypes[i][1] == 0) {
            typeMatch = true;
        } else if (type && wcsstr(type, g_config.targetTypes[i]) != nullptr) {
            typeMatch = true;
        }

        if (nameMatch && typeMatch) {
            // Heuristic: "Stroke" in the name means it's a stroke element
            if (name && wcsstr(name, L"Stroke") != nullptr) {
                *outIsStroke = true;
            }
            return true;
        }
    }
    return false;
}

// Discovery mode: log element for later analysis
void VisualTreeWatcher::LogElement(const VisualElement& element, InstanceHandle parent)
{
    if (!g_discoveryLog) return;

    const wchar_t* name = element.Name ? element.Name : L"(unnamed)";
    const wchar_t* type = element.Type ? element.Type : L"(unknown)";

    fprintf(g_discoveryLog, "[%llu] %ls | %ls (parent=%llu, numChildren=%u)\n",
        (unsigned long long)element.Handle,
        name, type,
        (unsigned long long)parent,
        element.NumChildren);
    fflush(g_discoveryLog);
}

// ── OnVisualTreeChange ──
HRESULT VisualTreeWatcher::OnVisualTreeChange(
    ParentChildRelation relation,
    VisualElement element,
    VisualMutationType mutationType)
{
    if (mutationType == Add) {
        // In discovery mode, log everything
        if (g_discoveryMode) {
            LogElement(element, relation.Parent);
        }

        // In targeting mode, check if this element matches a target
        if (!g_discoveryMode && element.Name && element.Type) {
            bool isStroke = false;
            if (MatchesTarget(element.Name, element.Type, &isStroke)) {
                DebugLog("MATCHED element: name='%ls' type='%ls' handle=%llu",
                    element.Name, element.Type, (unsigned long long)element.Handle);

                if (m_trackedCount < MAX_TRACKED) {
                    int slot = m_trackedCount++;
                    m_tracked[slot].handle = element.Handle;
                    wcsncpy_s(m_tracked[slot].name, element.Name, 63);
                    wcsncpy_s(m_tracked[slot].type, element.Type, 127);
                    m_tracked[slot].isStroke = isStroke;
                    m_tracked[slot].active = true;

                    // Apply current mode immediately
                    if (g_mode != MODE_DEFAULT) {
                        ApplyMode(g_mode);
                    }
                }
            }
        }
    }
    else if (mutationType == Remove) {
        // Remove tracked elements that were deleted
        for (int i = 0; i < m_trackedCount; i++) {
            if (m_tracked[i].handle == element.Handle) {
                m_tracked[i].active = false;
                m_tracked[i].handle = 0;
            }
        }
    }

    return S_OK;
}

HRESULT VisualTreeWatcher::OnElementStateChanged(
    InstanceHandle, VisualElementState, LPCWSTR)
{
    return S_OK;
}

// ── ApplyMode ──
void VisualTreeWatcher::ApplyMode(AppearanceMode mode)
{
    DebugLog("ApplyMode: mode=%d, trackedCount=%d", (int)mode, m_trackedCount);
    if (!m_pDiag) return;

    for (int i = 0; i < m_trackedCount; i++) {
        if (!m_tracked[i].active || m_tracked[i].handle == 0) continue;
        ApplyToElement(m_tracked[i].handle, mode, m_tracked[i].isStroke);
    }
}

// ── ApplyToElement via GetPropertyValuesChain + SetProperty ──
void VisualTreeWatcher::ApplyToElement(InstanceHandle handle,
                                        AppearanceMode mode,
                                        bool isStroke)
{
    IInspectable* pInspectable = nullptr;
    HRESULT hr = m_pDiag->GetIInspectableFromHandle(handle, &pInspectable);
    if (FAILED(hr) || !pInspectable) return;

    double opacity = 1.0;
    bool setFill = false;

    switch (mode) {
        case MODE_TRANSPARENT:
            opacity = 0.0;
            setFill = true;
            break;
        case MODE_ACRYLIC:
            opacity = isStroke ? 0.0 : 0.3;
            setFill = !isStroke;
            break;
        case MODE_DEFAULT:
            opacity = 1.0;
            setFill = false;
            break;
    }

    SetElementOpacity(pInspectable, opacity);
    pInspectable->Release();
}

// ── SetElementOpacity via GetPropertyValuesChain ──
void VisualTreeWatcher::SetElementOpacity(IInspectable* pElement, double opacity)
{
    if (!m_pService || !pElement) return;

    InstanceHandle handle = 0;
    HRESULT hr = m_pDiag->GetHandleFromIInspectable(pElement, &handle);
    if (FAILED(hr)) return;

    unsigned int propCount = 0;
    PropertyChainSource* pSources = nullptr;
    unsigned int srcCount = 0;
    PropertyChainValue* pValues = nullptr;

    hr = m_pService->GetPropertyValuesChain(handle, &srcCount, &pSources, &propCount, &pValues);
    if (FAILED(hr)) return;

    unsigned int fillIndex = UINT_MAX;
    unsigned int opacityIndex = UINT_MAX;

    for (unsigned int p = 0; p < propCount; p++) {
        if (pValues[p].PropertyName) {
            std::wstring propName(pValues[p].PropertyName);
            if (propName == L"Fill") {
                fillIndex = pValues[p].Index;
            }
            if (propName == L"Opacity") {
                opacityIndex = pValues[p].Index;
            }
        }
    }

    // Set Opacity
    if (opacityIndex != UINT_MAX) {
        std::wstring opValStr = std::to_wstring(opacity);
        InstanceHandle hValue = 0;
        BSTR bstrType = SysAllocString(L"Double");
        BSTR bstrVal = SysAllocString(opValStr.c_str());
        hr = m_pService->CreateInstance(bstrType, bstrVal, &hValue);
        SysFreeString(bstrType);
        SysFreeString(bstrVal);

        if (SUCCEEDED(hr)) {
            hr = m_pService->SetProperty(handle, hValue, opacityIndex);
            DebugLog("  SetProperty(opacity=%f, idx=%u) = 0x%08X", opacity, opacityIndex, hr);
        }
    }

    // Set Fill to transparent brush when making transparent
    if (fillIndex != UINT_MAX && opacity < 1.0) {
        InstanceHandle hBrush = 0;
        BSTR bstrType = SysAllocString(L"Windows.UI.Xaml.Media.SolidColorBrush");
        BSTR bstrVal = SysAllocString(L"Transparent");
        hr = m_pService->CreateInstance(bstrType, bstrVal, &hBrush);
        SysFreeString(bstrType);
        SysFreeString(bstrVal);

        if (SUCCEEDED(hr)) {
            hr = m_pService->SetProperty(handle, hBrush, fillIndex);
            DebugLog("  SetProperty(fill=Transparent, idx=%u) = 0x%08X", fillIndex, hr);
        }
    }

    // Free property chain
    for (unsigned int p = 0; p < propCount; p++) {
        if (pValues[p].PropertyName) SysFreeString(pValues[p].PropertyName);
        if (pValues[p].Value) SysFreeString(pValues[p].Value);
        if (pValues[p].Type) SysFreeString(pValues[p].Type);
        if (pValues[p].DeclaringType) SysFreeString(pValues[p].DeclaringType);
        if (pValues[p].ValueType) SysFreeString(pValues[p].ValueType);
        if (pValues[p].ItemType) SysFreeString(pValues[p].ItemType);
    }
    CoTaskMemFree(pValues);
    for (unsigned int s = 0; s < srcCount; s++) {
        if (pSources[s].Name) SysFreeString(pSources[s].Name);
        if (pSources[s].TargetType) SysFreeString(pSources[s].TargetType);
    }
    CoTaskMemFree(pSources);
}
