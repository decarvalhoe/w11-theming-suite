import useThemeStore from '../../store/themeStore';

export default function TopBar() {
  const themeName = useThemeStore((s) => s.theme.meta.name);
  const dirty = useThemeStore((s) => s.dirty);
  const loadTheme = useThemeStore((s) => s.loadTheme);
  const setLastAction = useThemeStore((s) => s.setLastAction);
  const setActiveSection = useThemeStore((s) => s.setActiveSection);

  const handleUndo = async () => {
    if (!window.themeAPI) return;
    try {
      const result = await window.themeAPI.restore('app-session');
      if (result.success) {
        const themeResult = await window.themeAPI.readTheme();
        if (themeResult.success) {
          loadTheme(themeResult.data);
        }
        setLastAction('Restored to session backup');
      }
    } catch (err) {
      setLastAction('Undo failed: ' + err.message, false);
    }
  };

  const handleSave = () => {
    setActiveSection('presets');
  };

  return (
    <div className="topbar">
      <div className="topbar__left">
        <span className="topbar__title">
          W11 Theme Configurator
        </span>
        <span style={{ color: '#666', fontSize: 12 }}>—</span>
        <span style={{ color: '#aaa', fontSize: 13 }}>{themeName}</span>
        {dirty && <span className="topbar__dirty">(unsaved)</span>}
      </div>
      <div className="topbar__right">
        <button className="btn" onClick={handleUndo}>↩ Undo All</button>
        <button className="btn btn--primary" onClick={handleSave}>💾 Save Preset</button>
      </div>
    </div>
  );
}
