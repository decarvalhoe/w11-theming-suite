# Test-LiveTransparency.ps1
# Applique la transparence et la LAISSE ACTIVE pour verification visuelle
$ErrorActionPreference = 'Stop'

Import-Module 'C:\Dev\w11-theming-suite\w11-theming-suite.psd1' -Force -DisableNameChecking

Write-Host "=== TRANSPARENCE LIVE ===" -ForegroundColor Cyan
Write-Host ""

# Check TTB is gone
$proc = Get-Process -Name 'TranslucentTB' -ErrorAction SilentlyContinue
if ($proc) {
    Write-Host "ATTENTION: TranslucentTB est actif!" -ForegroundColor Red
    exit 1
}
Write-Host "TranslucentTB: absent (OK)" -ForegroundColor Green

# Find all taskbar-related windows
$hwndMain = [W11ThemeSuite.TaskbarTransparency]::FindWindow('Shell_TrayWnd', $null)
$hwndXaml = [W11ThemeSuite.TaskbarTransparency]::FindWindow('Shell_TrayWnd', 'Taskbar')
$hwndNew  = [W11ThemeSuite.TaskbarTransparency]::FindWindow('Windows.UI.Composition.DesktopWindowContentBridge', $null)

Write-Host ""
Write-Host "Handles trouves:" -ForegroundColor Cyan
Write-Host "  Shell_TrayWnd (null):    0x$($hwndMain.ToString('X')) $(if ($hwndMain -ne [IntPtr]::Zero) {'FOUND'} else {'NOT FOUND'})"
Write-Host "  Shell_TrayWnd (Taskbar): 0x$($hwndXaml.ToString('X')) $(if ($hwndXaml -ne [IntPtr]::Zero) {'FOUND'} else {'NOT FOUND'})"
Write-Host "  DesktopWindowContent:    0x$($hwndNew.ToString('X'))  $(if ($hwndNew  -ne [IntPtr]::Zero) {'FOUND'} else {'NOT FOUND'})"

Write-Host ""
Write-Host "--- Test de TOUS les AccentState sur Shell_TrayWnd ---" -ForegroundColor Yellow

$states = @(
    @{ State = 0; Name = 'ACCENT_DISABLED (normal)' },
    @{ State = 1; Name = 'ACCENT_ENABLE_GRADIENT' },
    @{ State = 2; Name = 'ACCENT_ENABLE_TRANSPARENTGRADIENT (clear)' },
    @{ State = 3; Name = 'ACCENT_ENABLE_BLURBEHIND (blur)' },
    @{ State = 4; Name = 'ACCENT_ENABLE_ACRYLICBLURBEHIND (acrylic)' },
    @{ State = 5; Name = 'ACCENT_ENABLE_HOSTBACKDROP' }
)

foreach ($s in $states) {
    Write-Host ""
    Write-Host ">>> AccentState = $($s.State) : $($s.Name)" -ForegroundColor Magenta

    # Try on main taskbar
    $r1 = [W11ThemeSuite.TaskbarTransparency]::Apply($hwndMain, $s.State, [uint32]0)
    Write-Host "    Shell_TrayWnd: API retourne $r1"

    # Also try with AccentFlags variations
    # AccentFlags: 0, 2 (standard), 480 (some implementations use this)

    Write-Host "    REGARDE ta taskbar maintenant... (5s)" -ForegroundColor Yellow
    Start-Sleep -Seconds 5
}

# Now try AccentFlags = 0 instead of 2
Write-Host ""
Write-Host "--- Test avec AccentFlags = 0 (au lieu de 2) ---" -ForegroundColor Yellow

# We need to call the raw API for this
$testCode = @'
using System;
using System.Runtime.InteropServices;

namespace W11Test {
    [StructLayout(LayoutKind.Sequential)]
    public struct AP { public int S; public int F; public uint C; public int A; }

    [StructLayout(LayoutKind.Sequential)]
    public struct WD { public int At; public IntPtr D; public int Sz; }

    public static class T {
        [DllImport("user32.dll")]
        public static extern int SetWindowCompositionAttribute(IntPtr h, ref WD d);

        public static bool Apply(IntPtr h, int state, int flags, uint color) {
            var a = new AP { S = state, F = flags, C = color, A = 0 };
            var d = new WD { At = 19, Sz = Marshal.SizeOf(a) };
            var p = Marshal.AllocHGlobal(d.Sz);
            try {
                Marshal.StructureToPtr(a, p, false);
                d.D = p;
                return SetWindowCompositionAttribute(h, ref d) != 0;
            } finally { Marshal.FreeHGlobal(p); }
        }
    }
}
'@
Add-Type -TypeDefinition $testCode -ErrorAction SilentlyContinue

# Test different flag combinations with AccentState=2
$flagTests = @(0, 2, 480, 6)
foreach ($flag in $flagTests) {
    Write-Host ""
    Write-Host ">>> AccentState=2, AccentFlags=$flag" -ForegroundColor Magenta
    $r = [W11Test.T]::Apply($hwndMain, 2, $flag, [uint32]0)
    Write-Host "    Retour: $r"
    Write-Host "    REGARDE... (4s)" -ForegroundColor Yellow
    Start-Sleep -Seconds 4
}

# Leave it on clear for final check
Write-Host ""
Write-Host "=== Laisse en CLEAR (AccentState=2, Flags=2) ===" -ForegroundColor Green
[W11ThemeSuite.TaskbarTransparency]::Apply($hwndMain, 2, [uint32]0) | Out-Null
Write-Host "La taskbar RESTE transparente. Verifie visuellement." -ForegroundColor Green
Write-Host "Pour reset: Set-W11NativeTaskbarTransparency -Style normal" -ForegroundColor Gray
