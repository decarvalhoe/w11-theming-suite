import { useState, useEffect } from 'react';
import useThemeStore from '../../store/themeStore';

export default function PresetsSection() {
  const [presets, setPresets] = useState([]);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [saveName, setSaveName] = useState('');
  const theme = useThemeStore((s) => s.theme);
  const loadTheme = useThemeStore((s) => s.loadTheme);
  const setLastAction = useThemeStore((s) => s.setLastAction);
  const markClean = useThemeStore((s) => s.markClean);

  const fetchPresets = async () => {
    if (!window.themeAPI) return;
    setLoading(true);
    try {
      const result = await window.themeAPI.listPresets();
      if (result.success) setPresets(result.data || []);
      else setLastAction('Failed to load presets: ' + (result.error || 'Unknown error'), false);
    } catch (err) {
      console.error('Failed to fetch presets:', err);
      setLastAction('Failed to fetch presets: ' + err.message, false);
    }
    setLoading(false);
  };

  useEffect(() => { fetchPresets(); }, []);

  const handleLoad = async (preset) => {
    try {
      const fileName = (preset.FileName || preset.Name || '').replace(/\.json$/, '');
      if (!fileName) { setLastAction('Load error: preset has no filename', false); return; }
      const result = await window.themeAPI.loadPreset(fileName, preset.Source);
      if (result.success) {
        loadTheme(result.data);
        // Apply full theme
        await window.themeAPI.applyFullTheme(result.data);
        setLastAction(`Loaded preset: ${preset.Name}`);
      } else {
        setLastAction('Load error: ' + (result.error || 'Unknown error'), false);
      }
    } catch (err) {
      setLastAction('Load error: ' + err.message, false);
    }
  };

  const handleSave = async () => {
    if (!saveName.trim()) return;
    setSaving(true);
    try {
      const configToSave = {
        ...theme,
        meta: { ...theme.meta, name: saveName.trim() }
      };
      const result = await window.themeAPI.savePreset(
        saveName.trim().toLowerCase().replace(/\s+/g, '-'),
        configToSave
      );
      if (result.success) {
        setLastAction(`Saved preset: ${saveName.trim()}`);
        markClean();
        setSaveName('');
        fetchPresets();
      } else {
        setLastAction('Save failed: ' + (result.error || 'Unknown error'), false);
      }
    } catch (err) {
      setLastAction('Save error: ' + err.message, false);
    }
    setSaving(false);
  };

  return (
    <div>
      <h2 className="content__header">Presets</h2>
      <p className="content__desc">Load built-in or user presets, and save the current configuration.</p>

      {/* Save new preset */}
      <div className="setting-group">
        <div className="setting-group__title">Save Current as Preset</div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <input type="text" value={saveName} onChange={(e) => setSaveName(e.target.value)}
            placeholder="my-custom-theme"
            onKeyDown={(e) => { if (e.key === 'Enter' && saveName.trim()) handleSave(); }}
            style={{
              background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
              borderRadius: 'var(--radius-sm)', padding: '6px 12px', color: 'var(--text-primary)',
              fontSize: 13, flex: 1, maxWidth: 300
            }} />
          <button className="btn btn--primary" onClick={handleSave} disabled={!saveName.trim() || saving}>
            {saving ? 'Saving...' : 'Save'}
          </button>
        </div>
      </div>

      {/* Preset list */}
      <div className="setting-group">
        <div className="setting-group__title">
          Available Presets
          <button className="btn" onClick={fetchPresets} style={{ marginLeft: 10, fontSize: 10, padding: '2px 8px' }}>
            Refresh
          </button>
        </div>

        {loading && <p style={{ color: 'var(--text-muted)' }}>Loading...</p>}

        {!loading && presets.length === 0 && (
          <p style={{ color: 'var(--text-muted)' }}>No presets found.</p>
        )}

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 10, marginTop: 8 }}>
          {presets.map((preset) => (
            <div key={preset.FileName || preset.Name || Math.random()} style={{
              background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
              borderRadius: 'var(--radius)', padding: 14, cursor: 'pointer',
              transition: 'border-color 0.12s'
            }}
              onMouseEnter={(e) => e.currentTarget.style.borderColor = 'var(--accent)'}
              onMouseLeave={(e) => e.currentTarget.style.borderColor = 'var(--border-color)'}
              onClick={() => handleLoad(preset)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
                <span style={{ fontWeight: 600, fontSize: 14 }}>{preset.Name}</span>
                <span style={{
                  fontSize: 10, padding: '2px 6px', borderRadius: 3,
                  background: preset.Source === 'preset' ? '#1a3a5c' : '#3a2a1c',
                  color: preset.Source === 'preset' ? '#6ab0ff' : '#ffaa55'
                }}>
                  {preset.Source === 'preset' ? 'Built-in' : 'User'}
                </span>
              </div>
              <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                {preset.Description || 'No description'}
              </div>
              <div style={{ fontSize: 11, color: 'var(--text-muted)', marginTop: 4 }}>
                v{preset.Version} by {preset.Author}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
