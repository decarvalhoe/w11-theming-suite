import useThemeStore from './store/themeStore';
import { useThemeInit } from './hooks/useLivePreview';

import Sidebar from './components/layout/Sidebar';
import TopBar from './components/layout/TopBar';
import StatusBar from './components/layout/StatusBar';

import ModeSection from './components/sections/ModeSection';
import AccentColorSection from './components/sections/AccentColorSection';
import DwmSection from './components/sections/DwmSection';
import TaskbarSection from './components/sections/TaskbarSection';
import Win32ColorsSection from './components/sections/Win32ColorsSection';
import WallpaperSection from './components/sections/WallpaperSection';
import TransparencySection from './components/sections/TransparencySection';
import CursorsSection from './components/sections/CursorsSection';
import SoundsSection from './components/sections/SoundsSection';
import VisualStylesSection from './components/sections/VisualStylesSection';
import DesktopIconsSection from './components/sections/DesktopIconsSection';
import AdvancedSection from './components/sections/AdvancedSection';
import PresetsSection from './components/sections/PresetsSection';

const sectionComponents = {
  mode: ModeSection,
  accentColor: AccentColorSection,
  dwm: DwmSection,
  taskbar: TaskbarSection,
  win32Colors: Win32ColorsSection,
  wallpaper: WallpaperSection,
  transparency: TransparencySection,
  cursors: CursorsSection,
  sounds: SoundsSection,
  visualStyles: VisualStylesSection,
  desktopIcons: DesktopIconsSection,
  advanced: AdvancedSection,
  presets: PresetsSection,
};

export default function App() {
  // Load system theme state on first mount
  useThemeInit();

  const activeSection = useThemeStore((s) => s.activeSection);
  const ActiveComponent = sectionComponents[activeSection] || ModeSection;

  return (
    <div className="app-layout">
      <TopBar />
      <Sidebar />
      <main className="content">
        <ActiveComponent />
      </main>
      <StatusBar />
    </div>
  );
}
