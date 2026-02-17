// ShellTAP.h -- Generic Shell Transparency/Appearance Plugin for w11-theming-suite
// Refactored from TaskbarTAP to support injection into ANY XAML-based process:
//   - explorer.exe (taskbar)
//   - StartMenuExperienceHost.exe (Start Menu)
//   - ShellExperienceHost.exe (Action Center, Notifications)
//
// Configuration is passed via a named shared memory region whose name
// is derived from a TargetId (e.g., "Taskbar", "StartMenu", "ActionCenter").
//
// Based on techniques from RainbowTaskbar (MIT) and TranslucentTB (GPL).
// This implementation is original code for w11-theming-suite.
#pragma once

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <ocidl.h>     // IObjectWithSite
#include <unknwn.h>     // IClassFactory
#include <xamlOM.h>     // IVisualTreeService3, IXamlDiagnostics, IVisualTreeServiceCallback2
#include <oleauto.h>    // SysAllocString, SysFreeString
#include <atomic>

// Forward declarations
class ShellTAPSite;
class VisualTreeWatcher;

// ── Global state ──
extern HMODULE g_hModule;
extern std::atomic<long> g_refCount;
extern IVisualTreeService3* g_pTreeService;
extern IXamlDiagnostics* g_pDiagnostics;

// Appearance modes (same values as TaskbarTAP for compat)
enum AppearanceMode {
    MODE_DEFAULT     = 0,
    MODE_TRANSPARENT = 1,
    MODE_ACRYLIC     = 2
};

extern AppearanceMode g_mode;

// ── Shared memory configuration ──
// PowerShell writes this struct to "W11ThemeSuite_ShellTAP_<TargetId>"
// The DLL reads it on init and monitors for mode changes.
#pragma pack(push, 1)
struct ShellTAPConfig {
    int      version;            // Must be 1
    int      mode;               // 0=Default, 1=Transparent, 2=Acrylic
    int      targetCount;        // 0 = discovery mode (log ALL elements)
    wchar_t  targetNames[8][64]; // Element names to match (e.g., "BackgroundFill")
    wchar_t  targetTypes[8][128];// Element types to match (e.g., "Rectangle")
    wchar_t  logPath[260];       // Path for discovery log output
    int      flags;              // Reserved for future use
};
#pragma pack(pop)

static const int SHELLTAP_CONFIG_VERSION = 1;

// ── Exported C functions ──
extern "C" {
    __declspec(dllexport) HRESULT __stdcall SetShellTAPMode(int mode);
    __declspec(dllexport) int     __stdcall GetShellTAPMode();
    __declspec(dllexport) int     __stdcall GetShellTAPVersion();
    __declspec(dllexport) int     __stdcall GetShellTAPAppliedCount();
}

// ── COM class: ShellTAPSite -- receives XAML diagnostics site ──
class ShellTAPSite : public IObjectWithSite {
public:
    ShellTAPSite();
    virtual ~ShellTAPSite();

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;

    // IObjectWithSite
    HRESULT STDMETHODCALLTYPE SetSite(IUnknown* pUnkSite) override;
    HRESULT STDMETHODCALLTYPE GetSite(REFIID riid, void** ppvSite) override;

private:
    long m_refCount;
    IUnknown* m_pSite;
};

// ── COM class: VisualTreeWatcher -- watches XAML tree changes ──
class VisualTreeWatcher : public IVisualTreeServiceCallback2 {
public:
    VisualTreeWatcher(IXamlDiagnostics* pDiag, IVisualTreeService3* pService);
    virtual ~VisualTreeWatcher();

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;

    // IVisualTreeServiceCallback
    HRESULT STDMETHODCALLTYPE OnVisualTreeChange(
        ParentChildRelation relation,
        VisualElement element,
        VisualMutationType mutationType) override;

    // IVisualTreeServiceCallback2
    HRESULT STDMETHODCALLTYPE OnElementStateChanged(
        InstanceHandle element,
        VisualElementState elementState,
        LPCWSTR context) override;

    // Apply current mode to all tracked elements
    void ApplyMode(AppearanceMode mode);

    // Get count of tracked elements
    int GetTrackedCount() const { return m_trackedCount; }

private:
    // Set property via GetPropertyValuesChain + SetProperty
    void ApplyToElement(InstanceHandle handle, AppearanceMode mode, bool isStroke);
    void SetElementOpacity(IInspectable* pElement, double opacity);

    // Discovery: log all elements
    void LogElement(const VisualElement& element, InstanceHandle parent);

    long m_refCount;
    IXamlDiagnostics* m_pDiag;
    IVisualTreeService3* m_pService;

    // Tracked XAML element handles (matched from config targets)
    static const int MAX_TRACKED = 32;
    struct TrackedElement {
        InstanceHandle handle;
        wchar_t name[64];
        wchar_t type[128];
        bool isStroke;  // treat as stroke (set opacity=0 always when transparent)
        bool active;
    };
    TrackedElement m_tracked[MAX_TRACKED];
    int m_trackedCount;

    // Check if an element matches any configured target
    bool MatchesTarget(const wchar_t* name, const wchar_t* type, bool* outIsStroke);
};

// ── COM class factory ──
class ShellTAPFactory : public IClassFactory {
public:
    ShellTAPFactory();

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pOuter, REFIID riid, void** ppv) override;
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override;

private:
    long m_refCount;
};
