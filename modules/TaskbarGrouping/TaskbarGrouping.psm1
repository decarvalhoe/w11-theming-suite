Set-StrictMode -Version Latest

# ===========================================================================
# TaskbarGrouping.psm1
# ===========================================================================
# Force taskbar grouping of multiple instances of an application by profile,
# on Windows 11 (including build 26200+ where AUMID and window class alone
# are insufficient — Microsoft regressions in the new XAML taskbar).
#
# Three techniques in cascade, applied in order of effectiveness:
#
#   1) PER-PROFILE EXE HARDLINK
#      Create wezterm-gui-<profile>.exe as a hardlink to the original EXE.
#      Windows 11 26200 groups taskbar buttons by image path/inode of the
#      backing process, so distinct hardlink names = distinct taskbar groups.
#      This is the ONLY technique that reliably works on Win11 26200+.
#
#   2) WIN32 WINDOW CLASS OVERRIDE
#      Pass --class <NAME> to apps that support it (WezTerm). Windows uses
#      the window class to disambiguate apps with the same EXE — useful as
#      a complement when EXE hardlinks aren't an option.
#
#   3) PKEY_AppUserModel_ID
#      Set the per-window AUMID via SHGetPropertyStoreForWindow. Officially
#      the canonical Windows API for taskbar grouping, but Win11 26200 only
#      honors it inconsistently. Set as belt-and-suspenders.
#
# Designed primarily for terminal multiplexers (WezTerm, Windows Terminal,
# alacritty) but works for any EXE that supports being relaunched via a
# differently-named copy of itself.
#
# Validated on Windows 11 Pro 26200, Build 25H2.
#
# Public functions:
#   Set-W11TaskbarGrouping       — one-shot configuration per profile
#   Get-W11TaskbarGrouping       — list current grouping config
#   Remove-W11TaskbarGrouping    — remove aliases + AUMID for a profile
#   New-W11TaskbarExeAlias       — create just the EXE hardlink
#   Get-W11TaskbarExeAlias       — list aliases on disk
#   Remove-W11TaskbarExeAlias    — remove an EXE alias
#   Set-W11WindowAumid           — set AUMID on a HWND
#   Get-W11WindowAumid           — read AUMID from a HWND
#   Get-W11TaskbarGroupingLaunchSpec — returns the exec args a launcher
#                                       should use for a profile
# ===========================================================================

# ---------------------------------------------------------------------------
# P/Invoke: SHGetPropertyStoreForWindow + IPropertyStore + PKEY_AppUserModel_ID
# ---------------------------------------------------------------------------
if (-not ('W11TaskbarGrouping.AumidNative' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace W11TaskbarGrouping {
    public static class AumidNative {
        [DllImport("user32.dll", CharSet=CharSet.Auto, SetLastError=true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder buf, int maxLen);

        [DllImport("shell32.dll", PreserveSig=false)]
        public static extern void SHGetPropertyStoreForWindow(
            IntPtr hwnd,
            ref Guid iid,
            [MarshalAs(UnmanagedType.Interface)] out IPropertyStore ppv);

        [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [ComImport]
        public interface IPropertyStore {
            void GetCount(out uint cProps);
            void GetAt(uint iProp, out PROPERTYKEY pkey);
            void GetValue(ref PROPERTYKEY key, out PropVariant pv);
            void SetValue(ref PROPERTYKEY key, ref PropVariant pv);
            void Commit();
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROPERTYKEY { public Guid fmtid; public uint pid; }

        [StructLayout(LayoutKind.Sequential)]
        public struct PropVariant {
            public ushort vt;
            public ushort r1, r2, r3;
            public IntPtr pwszVal;
            public IntPtr p2;
        }

        public const ushort VT_LPWSTR = 31;

        // PKEY_AppUserModel_ID = {9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}, 5
        public static readonly Guid PKEY_AumidFmtid =
            new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");

        public static readonly Guid IID_IPropertyStore =
            new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");

        public static bool SetAumid(IntPtr hwnd, string aumid) {
            try {
                Guid iid = IID_IPropertyStore;
                IPropertyStore ps;
                SHGetPropertyStoreForWindow(hwnd, ref iid, out ps);
                PROPERTYKEY key = new PROPERTYKEY { fmtid = PKEY_AumidFmtid, pid = 5 };
                PropVariant pv = new PropVariant {
                    vt      = VT_LPWSTR,
                    pwszVal = Marshal.StringToCoTaskMemUni(aumid)
                };
                ps.SetValue(ref key, ref pv);
                ps.Commit();
                Marshal.FreeCoTaskMem(pv.pwszVal);
                Marshal.ReleaseComObject(ps);
                return true;
            } catch (Exception) { return false; }
        }

        public static string GetAumid(IntPtr hwnd) {
            try {
                Guid iid = IID_IPropertyStore;
                IPropertyStore ps;
                SHGetPropertyStoreForWindow(hwnd, ref iid, out ps);
                PROPERTYKEY key = new PROPERTYKEY { fmtid = PKEY_AumidFmtid, pid = 5 };
                PropVariant pv;
                ps.GetValue(ref key, out pv);
                string s = (pv.vt == VT_LPWSTR && pv.pwszVal != IntPtr.Zero)
                    ? Marshal.PtrToStringUni(pv.pwszVal) : null;
                Marshal.ReleaseComObject(ps);
                return s;
            } catch (Exception) { return null; }
        }

        public static string GetWindowClassName(IntPtr hwnd) {
            var sb = new StringBuilder(256);
            GetClassName(hwnd, sb, 256);
            return sb.ToString();
        }
    }
}
"@
}

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------
function script:Get-AliasRoot {
    $root = Join-Path $env:LOCALAPPDATA 'W11TaskbarGroupingAliases'
    if (-not (Test-Path $root)) {
        New-Item -ItemType Directory -Path $root -Force | Out-Null
    }
    return $root
}

function script:Get-AliasPath {
    param(
        [Parameter(Mandatory)] [string] $SourceExe,
        [Parameter(Mandatory)] [string] $Profile
    )
    $aliasRoot = Get-AliasRoot
    $base      = [System.IO.Path]::GetFileNameWithoutExtension($SourceExe)
    $ext       = [System.IO.Path]::GetExtension($SourceExe)
    return Join-Path $aliasRoot ("{0}-{1}{2}" -f $base, $Profile, $ext)
}

function script:Test-IsHardLinkable {
    param([string] $SourcePath, [string] $TargetPath)
    # Hardlinks must be on the SAME volume.
    $srcRoot = [System.IO.Path]::GetPathRoot($SourcePath).TrimEnd('\')
    $tgtRoot = [System.IO.Path]::GetPathRoot($TargetPath).TrimEnd('\')
    return ($srcRoot -ieq $tgtRoot)
}

function script:Copy-RuntimeDependencies {
    # Some EXEs (WezTerm, Electron apps) need sibling DLLs to load. When we
    # alias into a different folder, copy the sibling files so the alias is
    # self-contained. Hardlinks where possible to save disk.
    param([string] $SourceExe, [string] $TargetDir)

    $srcDir = Split-Path -Parent $SourceExe
    Get-ChildItem $srcDir -File -Include *.dll,*.exe -Recurse:$false |
        Where-Object { $_.FullName -ne $SourceExe } |
        ForEach-Object {
            $dst = Join-Path $TargetDir $_.Name
            if (Test-Path $dst) { return }
            if (Test-IsHardLinkable -SourcePath $_.FullName -TargetPath $dst) {
                try {
                    New-Item -ItemType HardLink -Path $dst -Target $_.FullName -ErrorAction Stop | Out-Null
                } catch {
                    Copy-Item $_.FullName $dst -Force -ErrorAction SilentlyContinue
                }
            } else {
                Copy-Item $_.FullName $dst -Force -ErrorAction SilentlyContinue
            }
        }
}

# ---------------------------------------------------------------------------
# EXE alias management
# ---------------------------------------------------------------------------
function New-W11TaskbarExeAlias {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string] $SourceExe,
        [Parameter(Mandatory)] [string] $Profile,
        [switch] $WithDependencies
    )

    if (-not (Test-Path $SourceExe)) {
        throw "SourceExe not found: $SourceExe"
    }

    $aliasPath = Get-AliasPath -SourceExe $SourceExe -Profile $Profile
    $aliasDir  = Split-Path -Parent $aliasPath

    if ($PSCmdlet.ShouldProcess($aliasPath, "Create EXE alias for profile '$Profile'")) {
        if (Test-Path $aliasPath) {
            Remove-Item $aliasPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-IsHardLinkable -SourcePath $SourceExe -TargetPath $aliasPath) {
            try {
                New-Item -ItemType HardLink -Path $aliasPath -Target $SourceExe -ErrorAction Stop | Out-Null
                Write-Verbose "Hardlink created: $aliasPath -> $SourceExe"
            } catch {
                Copy-Item $SourceExe $aliasPath -Force
                Write-Verbose "Fallback copy: $aliasPath"
            }
        } else {
            Copy-Item $SourceExe $aliasPath -Force
            Write-Verbose "Cross-volume, copy: $aliasPath"
        }

        if ($WithDependencies) {
            Copy-RuntimeDependencies -SourceExe $SourceExe -TargetDir $aliasDir
        }
    }

    return [pscustomobject]@{
        Profile     = $Profile
        SourceExe   = $SourceExe
        AliasPath   = $aliasPath
        AliasName   = [System.IO.Path]::GetFileNameWithoutExtension($aliasPath)
        Exists      = Test-Path $aliasPath
        WithDeps    = [bool]$WithDependencies
    }
}

function Get-W11TaskbarExeAlias {
    [CmdletBinding()]
    param(
        [string] $Profile
    )
    $root = Get-AliasRoot
    if (-not (Test-Path $root)) { return @() }
    Get-ChildItem $root -File -Filter '*.exe' -ErrorAction SilentlyContinue |
        Where-Object { (-not $Profile) -or $_.Name -like "*-$Profile.exe" } |
        ForEach-Object {
            $base = $_.BaseName
            # Pattern: <name>-<profile>
            $idx = $base.LastIndexOf('-')
            $detectedProfile = if ($idx -gt 0) { $base.Substring($idx+1) } else { '' }
            [pscustomobject]@{
                AliasPath = $_.FullName
                AliasName = $base
                Profile   = $detectedProfile
                SizeKB    = [math]::Round($_.Length/1KB, 1)
                Modified  = $_.LastWriteTime
            }
        }
}

function Remove-W11TaskbarExeAlias {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string] $Profile,
        [string] $SourceExeName  # e.g. 'wezterm-gui' to scope removal
    )
    $root = Get-AliasRoot
    if (-not (Test-Path $root)) { return }
    Get-ChildItem $root -File -Filter '*.exe' -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like "*-$Profile.exe" -and
            ($null -eq $SourceExeName -or $_.Name -like "$SourceExeName-*")
        } |
        ForEach-Object {
            if ($PSCmdlet.ShouldProcess($_.FullName, 'Remove alias')) {
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
                Write-Verbose "Removed: $($_.FullName)"
            }
        }
}

# ---------------------------------------------------------------------------
# AUMID (per-window) management
# ---------------------------------------------------------------------------
function Set-W11WindowAumid {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)] [IntPtr] $WindowHandle,
        [Parameter(Mandatory)] [string] $Aumid
    )
    process {
        $ok = [W11TaskbarGrouping.AumidNative]::SetAumid($WindowHandle, $Aumid)
        [pscustomobject]@{
            WindowHandle = $WindowHandle
            Aumid        = $Aumid
            Success      = $ok
        }
    }
}

function Get-W11WindowAumid {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)] [IntPtr] $WindowHandle
    )
    process {
        $aumid = [W11TaskbarGrouping.AumidNative]::GetAumid($WindowHandle)
        $klass = [W11TaskbarGrouping.AumidNative]::GetWindowClassName($WindowHandle)
        [pscustomobject]@{
            WindowHandle = $WindowHandle
            Aumid        = $aumid
            WindowClass  = $klass
        }
    }
}

# ---------------------------------------------------------------------------
# High-level configuration
# ---------------------------------------------------------------------------
function Set-W11TaskbarGrouping {
    <#
    .SYNOPSIS
        Configure taskbar grouping for a profile.

    .DESCRIPTION
        Creates a per-profile EXE hardlink (the only reliable Win11 26200
        technique) and exposes a launch spec containing the recommended
        argument list to spawn instances that will be grouped together in
        the taskbar.

        Optionally also forces the global "Combine taskbar buttons" registry
        setting to "Always combine" (TaskbarGlomLevel=0, MMTaskbarGlomLevel=0)
        and restarts explorer so the change takes effect.

    .EXAMPLE
        Set-W11TaskbarGrouping -SourceExe 'C:\Program Files\WezTerm\wezterm-gui.exe' `
                              -Profile 'rbok' -WithDependencies

        Creates wezterm-gui-rbok.exe + sibling DLLs in
        $env:LOCALAPPDATA\W11TaskbarGroupingAliases\ and returns a launch
        spec the caller can use to spawn windows.

    .EXAMPLE
        $spec = Set-W11TaskbarGrouping -SourceExe 'wezterm-gui.exe' -Profile 'rbok'
        Start-Process -FilePath $spec.AliasPath -ArgumentList $spec.RecommendedClassArg, "start", "--"
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string] $SourceExe,
        [Parameter(Mandatory)] [string] $Profile,
        [string]              $AumidPrefix = 'W11ThemingSuite.TaskbarGroup',
        [switch]              $WithDependencies,
        [switch]              $ForceCombineRegistry,
        [switch]              $RestartExplorer
    )

    # 1) EXE alias (most reliable on Win11 26200+)
    $alias = New-W11TaskbarExeAlias -SourceExe $SourceExe -Profile $Profile -WithDependencies:$WithDependencies

    # 2) Optional: force "Always combine" in registry
    if ($ForceCombineRegistry) {
        $advKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
        Set-ItemProperty -Path $advKey -Name TaskbarGlomLevel    -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $advKey -Name MMTaskbarGlomLevel  -Value 0 -Type DWord -Force
        Write-Verbose 'TaskbarGlomLevel = 0 (Always combine) applied.'
    }

    # 3) Optional: restart explorer
    if ($RestartExplorer) {
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        Start-Process explorer.exe
    }

    return [pscustomobject]@{
        Profile         = $Profile
        SourceExe       = $SourceExe
        AliasPath       = $alias.AliasPath
        AliasName       = $alias.AliasName
        Aumid           = "$AumidPrefix.$Profile"
        WithDeps        = $alias.WithDeps
        Status          = if ($alias.Exists) { 'Ready' } else { 'Failed' }
    }
}

function Get-W11TaskbarGrouping {
    [CmdletBinding()]
    param([string] $Profile)

    $aliases = Get-W11TaskbarExeAlias -Profile $Profile
    if (-not $aliases) { return @() }

    $aliases | ForEach-Object {
        $aliasNameOnly = [System.IO.Path]::GetFileNameWithoutExtension($_.AliasName)
        $procs = Get-Process -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -ieq $aliasNameOnly -and $_.MainWindowHandle -ne 0 }

        $windows = @($procs | ForEach-Object {
            [pscustomobject]@{
                PID    = $_.Id
                HWND   = $_.MainWindowHandle
                Title  = $_.MainWindowTitle
                Class  = [W11TaskbarGrouping.AumidNative]::GetWindowClassName($_.MainWindowHandle)
                Aumid  = [W11TaskbarGrouping.AumidNative]::GetAumid($_.MainWindowHandle)
            }
        })

        [pscustomobject]@{
            Profile      = $_.Profile
            AliasPath    = $_.AliasPath
            AliasName    = $_.AliasName
            WindowsCount = $windows.Count
            Windows      = $windows
        }
    }
}

function Remove-W11TaskbarGrouping {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string] $Profile,
        [string]              $SourceExeName,
        [switch]              $KillRunningWindows
    )

    if ($KillRunningWindows) {
        $aliasNameLike = if ($SourceExeName) { "$SourceExeName-$Profile" } else { "*-$Profile" }
        Get-Process -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like $aliasNameLike } |
            ForEach-Object {
                if ($PSCmdlet.ShouldProcess("PID $($_.Id)", 'Stop process')) {
                    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                }
            }
    }

    Remove-W11TaskbarExeAlias -Profile $Profile -SourceExeName $SourceExeName
}

# ---------------------------------------------------------------------------
# Launcher integration helper
# ---------------------------------------------------------------------------
function Get-W11TaskbarGroupingLaunchSpec {
    <#
    .SYNOPSIS
        Returns the exact launch spec a downstream launcher (e.g. the
        universal-project-launcher's wezterm_monitor.ps1) should use to
        spawn an instance with grouping applied.

    .EXAMPLE
        $spec = Get-W11TaskbarGroupingLaunchSpec -SourceExe (Get-Command wezterm-gui.exe).Path `
                                                  -Profile 'rbok' -App 'wezterm'
        Start-Process -FilePath $spec.ExecutablePath -ArgumentList $spec.ExtraArgs + @('start','--')
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SourceExe,
        [Parameter(Mandatory)] [string] $Profile,
        [ValidateSet('wezterm','generic')]
        [string]               $App = 'generic'
    )

    $aliasPath = Get-AliasPath -SourceExe $SourceExe -Profile $Profile
    if (-not (Test-Path $aliasPath)) {
        Write-Warning "Alias not found at $aliasPath. Call Set-W11TaskbarGrouping first."
    }

    # App-specific extra args
    $extraArgs = @()
    switch ($App) {
        'wezterm' {
            # WezTerm supports --class as a 'start' subcommand option.
            # Pass it so the Win32 window class also encodes the profile.
            $extraArgs = @('start', '--class', "W11ThemingSuite.$Profile", '--always-new-process')
        }
    }

    return [pscustomobject]@{
        ExecutablePath = if (Test-Path $aliasPath) { $aliasPath } else { $SourceExe }
        Profile        = $Profile
        Aumid          = "W11ThemingSuite.TaskbarGroup.$Profile"
        WindowClass    = "W11ThemingSuite.$Profile"
        ExtraArgs      = $extraArgs
        UsingAlias     = (Test-Path $aliasPath)
    }
}

# ===========================================================================
# Group focus management
# ===========================================================================
# Once windows are taskbar-grouped, clicking the group button only shows the
# thumbnail strip — it doesn't bring all 6 windows to the foreground at once.
# These helpers do exactly that: focus all windows of a profile in one shot.
#
# Useful for "switch context": Ctrl+Alt+R brings up all RBOK windows, then
# Ctrl+Alt+N brings up all NOMOS windows over them.
# ===========================================================================

if (-not ('W11TaskbarGrouping.WindowActions' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace W11TaskbarGrouping {
    public static class WindowActions {
        [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        public const int SW_RESTORE     = 9;
        public const int SW_SHOW        = 5;
        public const int SW_SHOWNOACTIVATE = 4;
        public static readonly IntPtr HWND_TOP    = new IntPtr(0);
        public static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        public const uint SWP_NOSIZE     = 0x0001;
        public const uint SWP_NOMOVE     = 0x0002;
        public const uint SWP_SHOWWINDOW = 0x0040;
    }
}
"@
}

function Show-W11TaskbarGroup {
    <#
    .SYNOPSIS
        Bring all windows of a profile to the foreground at once.

    .DESCRIPTION
        For each window of the profile (resolved via the per-profile EXE
        alias hardlink), restores it from minimized state and raises it to
        the top of the Z-order. The LAST raised window receives focus.

        Useful when 6 instances are taskbar-grouped (e.g. via
        ExplorerPatcher) and clicking the group only shows thumbnails — this
        function brings ALL 6 to front in a single action.

    .PARAMETER Profile
        Name of the grouping profile (e.g. 'rbok', 'nomos', '42t').

    .PARAMETER FocusFirst
        Activate the FIRST window after raising all (default: last raised
        gets focus, which is whichever the OS returns last in process
        enumeration order).

    .EXAMPLE
        Show-W11TaskbarGroup -Profile rbok

    .EXAMPLE
        # Bound to a global hotkey in the system tray launcher:
        Show-W11TaskbarGroup -Profile nomos
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Profile,
        [switch]              $FocusFirst
    )

    $alias = Get-W11TaskbarExeAlias -Profile $Profile | Select-Object -First 1
    if (-not $alias) {
        Write-Warning "No alias found for profile '$Profile'. Call Set-W11TaskbarGrouping first."
        return
    }
    $aliasNameOnly = [System.IO.Path]::GetFileNameWithoutExtension($alias.AliasName)

    $procs = Get-Process | Where-Object {
        $_.Name -ieq $aliasNameOnly -and $_.MainWindowHandle -ne [IntPtr]::Zero
    } | Sort-Object Id

    if (-not $procs) {
        Write-Warning "No running windows found for profile '$Profile' (alias: $aliasNameOnly)."
        return
    }

    $firstHwnd = $null
    foreach ($p in $procs) {
        $h = $p.MainWindowHandle
        if ($null -eq $firstHwnd) { $firstHwnd = $h }
        if ([W11TaskbarGrouping.WindowActions]::IsIconic($h)) {
            [W11TaskbarGrouping.WindowActions]::ShowWindowAsync($h, [W11TaskbarGrouping.WindowActions]::SW_RESTORE) | Out-Null
        }
        # SWP_SHOWWINDOW + HWND_TOP without changing size/position: pure Z-order raise
        [W11TaskbarGrouping.WindowActions]::SetWindowPos(
            $h,
            [W11TaskbarGrouping.WindowActions]::HWND_TOP,
            0, 0, 0, 0,
            [W11TaskbarGrouping.WindowActions]::SWP_NOSIZE -bor
            [W11TaskbarGrouping.WindowActions]::SWP_NOMOVE -bor
            [W11TaskbarGrouping.WindowActions]::SWP_SHOWWINDOW
        ) | Out-Null
    }

    # SetForegroundWindow only works reliably on a window owned by the
    # current foreground thread. We attach to the foreground thread first.
    $hwndToFocus = if ($FocusFirst) { $firstHwnd } else { $procs[-1].MainWindowHandle }
    [W11TaskbarGrouping.WindowActions]::SetForegroundWindow($hwndToFocus) | Out-Null

    return [pscustomobject]@{
        Profile     = $Profile
        AliasName   = $alias.AliasName
        WindowsCount= $procs.Count
        FocusedHwnd = $hwndToFocus
    }
}

# ===========================================================================
# ExplorerPatcher integration — opt-in escape hatch for Win11 26200+ where
# Microsoft regressed the XAML taskbar's grouping logic to the point where
# AUMID, window class AND distinct EXE names are all ignored.
#
# This is the ONLY exception in this project to the "no third-party software"
# rule. It is opt-in only (the user must explicitly call
# Install-W11ExplorerPatcherHelper). Once installed, ExplorerPatcher restores
# the Windows 10 taskbar that honors all the native grouping mechanisms
# implemented above.
#
# Project: https://github.com/valinet/ExplorerPatcher (MIT-licensed)
# ===========================================================================

function Test-W11ExplorerPatcherInstalled {
    <#
    .SYNOPSIS
        Detect whether ExplorerPatcher is installed on this machine.

    .DESCRIPTION
        Checks three independent signals: the registry key created at install,
        the dxgi.dll hook in System32 left by EP, and a winget upgrade query.
        Returns a structured object with version + signals so callers can
        decide whether to install / upgrade.
    #>
    [CmdletBinding()]
    param()

    $regKey      = 'HKCU:\Software\ExplorerPatcher'
    $hasReg      = Test-Path $regKey
    $regVersion  = if ($hasReg) {
        $rv = Get-ItemProperty $regKey -ErrorAction SilentlyContinue
        if ($rv -and ($rv.PSObject.Properties.Name -contains 'Version')) { $rv.Version } else { $null }
    } else { $null }
    $dxgiPath    = "$env:SystemRoot\System32\dxgi.dll.local\ep_taskbar.2.dll"
    $epDllPath   = "$env:SystemRoot\dxgi.dll"
    $hasDll      = (Test-Path $dxgiPath) -or (Test-Path $epDllPath)
    $wingetVer   = $null
    try {
        $w = winget list --id valinet.ExplorerPatcher --exact 2>$null | Select-String -Pattern 'valinet\.ExplorerPatcher' | Select-Object -First 1
        if ($w) {
            $parts = ($w.Line -split '\s{2,}') | Where-Object { $_ }
            if ($parts.Count -ge 3) { $wingetVer = $parts[2] }
        }
    } catch {}

    return [pscustomobject]@{
        Installed       = ($hasReg -or $hasDll -or [bool]$wingetVer)
        RegistryPresent = $hasReg
        DllPresent      = $hasDll
        RegistryVersion = $regVersion
        WingetVersion   = $wingetVer
    }
}

function Install-W11ExplorerPatcherHelper {
    <#
    .SYNOPSIS
        Install ExplorerPatcher (opt-in third-party escape hatch).

    .DESCRIPTION
        Tries winget first (silent install). Falls back to downloading
        ep_setup.exe from the latest GitHub release and launching the
        installer (interactive).

        After install, optionally calls Set-W11ExplorerPatcherTaskbarGrouping
        to switch the taskbar to "Windows 10 style" where grouping works.

    .PARAMETER UseWinget
        Force winget path. Default: try winget, fall back to GitHub.

    .PARAMETER Configure
        After install, call Set-W11ExplorerPatcherTaskbarGrouping to enable
        the Win10 taskbar with grouping = "Always combine".

    .EXAMPLE
        Install-W11ExplorerPatcherHelper -Configure
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch] $UseWinget,
        [switch] $Configure
    )

    $existing = Test-W11ExplorerPatcherInstalled
    if ($existing.Installed) {
        Write-Host "ExplorerPatcher already installed (registry=$($existing.RegistryPresent), dll=$($existing.DllPresent), winget=$($existing.WingetVersion))."
        if (-not $Configure) { return $existing }
    }

    if ($PSCmdlet.ShouldProcess('ExplorerPatcher', 'Install')) {
        $wingetOk = $false
        try {
            $wgPath = Get-Command winget -ErrorAction SilentlyContinue
            if ($wgPath) {
                Write-Host 'Installing via winget...'
                $args = @('install','--id','valinet.ExplorerPatcher','--exact','--accept-package-agreements','--accept-source-agreements','--silent')
                $p = Start-Process winget -ArgumentList $args -Wait -PassThru -NoNewWindow
                if ($p.ExitCode -eq 0) { $wingetOk = $true }
                else { Write-Warning "winget exited with code $($p.ExitCode); will try GitHub release fallback." }
            }
        } catch {
            Write-Warning "winget install failed: $_"
        }

        if (-not $wingetOk) {
            Write-Host 'Downloading latest ep_setup.exe from GitHub release...'
            $relJson = & gh api repos/valinet/ExplorerPatcher/releases/latest 2>$null
            if (-not $relJson) {
                Write-Warning 'gh CLI not available or release fetch failed. Aborting install.'
                return Test-W11ExplorerPatcherInstalled
            }
            $rel = $relJson | ConvertFrom-Json
            $asset = $rel.assets | Where-Object { $_.name -eq 'ep_setup.exe' } | Select-Object -First 1
            if (-not $asset) {
                Write-Warning 'No ep_setup.exe asset found in latest release. Aborting.'
                return Test-W11ExplorerPatcherInstalled
            }
            $tmpExe = Join-Path $env:TEMP 'ep_setup.exe'
            Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $tmpExe -UseBasicParsing
            Write-Host "Launching $tmpExe (UAC prompt expected)..."
            Start-Process $tmpExe -Verb RunAs -Wait
        }
    }

    $after = Test-W11ExplorerPatcherInstalled
    if ($Configure -and $after.Installed) {
        Set-W11ExplorerPatcherTaskbarGrouping -Style 'Windows10' -Combine 'Always' -RestartExplorer | Out-Null
    }
    return $after
}

function Set-W11ExplorerPatcherTaskbarGrouping {
    <#
    .SYNOPSIS
        Configure ExplorerPatcher for reliable taskbar grouping.

    .DESCRIPTION
        Writes the registry keys ExplorerPatcher reads at startup:
          - Taskbar_Style: 0 = Windows 10 style (grouping works), 1 = Win11
          - TaskbarGlomming / TaskbarGlomLevel as documented by EP
        Optionally restarts explorer so the change takes effect.

    .PARAMETER Style
        'Windows10' (recommended for grouping), 'Windows11' or 'Auto'.

    .PARAMETER Combine
        'Always' (default), 'WhenFull' or 'Never'.

    .EXAMPLE
        Set-W11ExplorerPatcherTaskbarGrouping -Style Windows10 -Combine Always -RestartExplorer
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [ValidateSet('Windows10','Windows11','Auto')]
        [string] $Style          = 'Windows10',

        [ValidateSet('Always','WhenFull','Never')]
        [string] $Combine        = 'Always',

        [switch] $RestartExplorer
    )

    if (-not (Test-W11ExplorerPatcherInstalled).Installed) {
        Write-Warning 'ExplorerPatcher is not installed. Call Install-W11ExplorerPatcherHelper first.'
        return
    }

    $regKey = 'HKCU:\Software\ExplorerPatcher'
    if (-not (Test-Path $regKey)) {
        New-Item -Path $regKey -Force | Out-Null
    }

    $styleVal = switch ($Style) { 'Windows10' { 0 } 'Windows11' { 1 } 'Auto' { 2 } }
    $combVal  = switch ($Combine) { 'Always' { 0 } 'WhenFull' { 1 } 'Never' { 2 } }

    if ($PSCmdlet.ShouldProcess('ExplorerPatcher', "Style=$Style, Combine=$Combine")) {
        Set-ItemProperty -Path $regKey -Name 'Taskbar_Style'     -Value $styleVal -Type DWord -Force
        Set-ItemProperty -Path $regKey -Name 'TaskbarGlomLevel'  -Value $combVal  -Type DWord -Force
        # Also keep the system-wide TaskbarGlomLevel in sync
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name TaskbarGlomLevel    -Value $combVal -Type DWord -Force
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name MMTaskbarGlomLevel  -Value $combVal -Type DWord -Force
    }

    if ($RestartExplorer) {
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        Start-Process explorer.exe
    }

    return [pscustomobject]@{
        Style    = $Style
        Combine  = $Combine
        Applied  = $true
        Restart  = [bool]$RestartExplorer
    }
}

Export-ModuleMember -Function @(
    'Set-W11TaskbarGrouping',
    'Get-W11TaskbarGrouping',
    'Remove-W11TaskbarGrouping',
    'New-W11TaskbarExeAlias',
    'Get-W11TaskbarExeAlias',
    'Remove-W11TaskbarExeAlias',
    'Set-W11WindowAumid',
    'Get-W11WindowAumid',
    'Get-W11TaskbarGroupingLaunchSpec',
    # Group focus management (bring all 6 windows to front in one shot)
    'Show-W11TaskbarGroup',
    # Opt-in ExplorerPatcher integration (only third-party escape hatch)
    'Test-W11ExplorerPatcherInstalled',
    'Install-W11ExplorerPatcherHelper',
    'Set-W11ExplorerPatcherTaskbarGrouping'
)
