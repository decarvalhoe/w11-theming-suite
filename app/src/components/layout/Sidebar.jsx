import useThemeStore from '../../store/themeStore';

const sections = [
  { id: 'mode',         icon: '🌗', label: 'Mode' },
  { id: 'accentColor',  icon: '🎨', label: 'Accent Color' },
  { id: 'dwm',          icon: '🪟', label: 'DWM' },
  { id: 'taskbar',      icon: '📌', label: 'Taskbar' },
  { id: 'win32Colors',  icon: '🖌️', label: 'Win32 Colors' },
  { id: 'separator1' },
  { id: 'wallpaper',    icon: '🖼️', label: 'Wallpaper' },
  { id: 'transparency', icon: '💎', label: 'Transparency' },
  { id: 'separator2' },
  { id: 'cursors',      icon: '🖱️', label: 'Cursors' },
  { id: 'sounds',       icon: '🔊', label: 'Sounds' },
  { id: 'visualStyles', icon: '🎭', label: 'Visual Styles' },
  { id: 'desktopIcons', icon: '🗂️', label: 'Desktop Icons' },
  { id: 'separator3' },
  { id: 'advanced',     icon: '⚙️', label: 'Advanced' },
  { id: 'presets',      icon: '📦', label: 'Presets' },
];

export default function Sidebar() {
  const activeSection = useThemeStore((s) => s.activeSection);
  const setActiveSection = useThemeStore((s) => s.setActiveSection);

  return (
    <nav className="sidebar">
      {sections.map((s) => {
        if (s.id.startsWith('separator')) {
          return <div key={s.id} className="sidebar__separator" />;
        }
        const isActive = activeSection === s.id;
        return (
          <div
            key={s.id}
            className={`sidebar__item ${isActive ? 'sidebar__item--active' : ''}`}
            onClick={() => setActiveSection(s.id)}
          >
            <span className="sidebar__icon">{s.icon}</span>
            <span>{s.label}</span>
          </div>
        );
      })}
    </nav>
  );
}
