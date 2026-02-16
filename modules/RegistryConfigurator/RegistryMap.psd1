# RegistryMap.psd1
# Maps logical theme section names to their Windows 11 registry paths.
# Each entry defines: Path (registry key), Name (value name), Type (registry value type).

@{
    DarkMode = @{
        AppsUseLightTheme = @{
            Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
            Name = 'AppsUseLightTheme'
            Type = 'DWord'
        }
        SystemUsesLightTheme = @{
            Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
            Name = 'SystemUsesLightTheme'
            Type = 'DWord'
        }
    }

    AccentColor = @{
        AccentColor = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'AccentColor'
            Type = 'DWord'
        }
        AccentColorMenu = @{
            Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent'
            Name = 'AccentColorMenu'
            Type = 'DWord'
        }
        StartColorMenu = @{
            Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent'
            Name = 'StartColorMenu'
            Type = 'DWord'
        }
        ColorPrevalence_Personalize = @{
            Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
            Name = 'ColorPrevalence'
            Type = 'DWord'
        }
        ColorPrevalence_DWM = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'ColorPrevalence'
            Type = 'DWord'
        }
        EnableTransparency = @{
            Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
            Name = 'EnableTransparency'
            Type = 'DWord'
        }
        AutoColorization = @{
            Path = 'HKCU:\Control Panel\Desktop'
            Name = 'AutoColorization'
            Type = 'DWord'
        }
    }

    DWM = @{
        ColorizationColor = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'ColorizationColor'
            Type = 'DWord'
        }
        ColorizationAfterglow = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'ColorizationAfterglow'
            Type = 'DWord'
        }
        ColorizationColorBalance = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'ColorizationColorBalance'
            Type = 'DWord'
        }
        ColorizationAfterglowBalance = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'ColorizationAfterglowBalance'
            Type = 'DWord'
        }
        ColorizationBlurBalance = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'ColorizationBlurBalance'
            Type = 'DWord'
        }
        ColorizationGlassReflectionIntensity = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'ColorizationGlassReflectionIntensity'
            Type = 'DWord'
        }
        EnableAeroPeek = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'EnableAeroPeek'
            Type = 'DWord'
        }
        ForceEffectMode = @{
            Path = 'HKCU:\Software\Microsoft\Windows\DWM'
            Name = 'ForceEffectMode'
            Type = 'DWord'
        }
    }

    Taskbar = @{
        TaskbarAl = @{
            Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
            Name = 'TaskbarAl'
            Type = 'DWord'
        }
        ShowTaskViewButton = @{
            Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
            Name = 'ShowTaskViewButton'
            Type = 'DWord'
        }
        TaskbarDa = @{
            Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
            Name = 'TaskbarDa'
            Type = 'DWord'
        }
        TaskbarMn = @{
            Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
            Name = 'TaskbarMn'
            Type = 'DWord'
        }
        SearchboxTaskbarMode = @{
            Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'
            Name = 'SearchboxTaskbarMode'
            Type = 'DWord'
        }
        UseOLEDTaskbarTransparency = @{
            Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
            Name = 'UseOLEDTaskbarTransparency'
            Type = 'DWord'
        }
    }
}
