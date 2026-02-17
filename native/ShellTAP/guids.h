// guids.h -- CLSIDs for ShellTAP (generic XAML injection)
// Must match what PowerShell passes to InitializeXamlDiagnosticsEx
#pragma once
#include <guiddef.h>

// {A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
DEFINE_GUID(CLSID_ShellTAPSite,
    0xa1b2c3d4, 0xe5f6, 0x7890,
    0xab, 0xcd, 0xef, 0x12, 0x34, 0x56, 0x78, 0x90);
