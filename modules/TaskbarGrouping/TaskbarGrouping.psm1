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

Export-ModuleMember -Function @(
    'Set-W11TaskbarGrouping',
    'Get-W11TaskbarGrouping',
    'Remove-W11TaskbarGrouping',
    'New-W11TaskbarExeAlias',
    'Get-W11TaskbarExeAlias',
    'Remove-W11TaskbarExeAlias',
    'Set-W11WindowAumid',
    'Get-W11WindowAumid',
    'Get-W11TaskbarGroupingLaunchSpec'
)
