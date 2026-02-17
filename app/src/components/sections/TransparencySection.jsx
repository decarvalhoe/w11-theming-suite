import useThemeStore from '../../store/themeStore';
import Toggle from '../controls/Toggle';
import Select from '../controls/Select';
import ColorPicker from '../controls/ColorPicker';

export default function TransparencySection() {
  const transparency = useThemeStore((s) => s.theme.transparency);
  const setField = useThemeStore((s) => s.setField);
  const setLastAction = useThemeStore((s) => s.setLastAction);

  const updateSub = (sub, key, value) => {
    const current = transparency[sub] || {};
    setField('transparency', sub, { ...current, [key]: value });
  };

  const apply = async (subsystem) => {
    if (!window.themeAPI) return;
    try {
      const params = transparency[subsystem];
      await window.themeAPI.applyTransparency(subsystem, params);
      setLastAction(`Applied ${subsystem} transparency`);
    } catch (err) {
      setLastAction(`Error: ${err.message}`, false);
    }
  };

  const tb = transparency.taskbar || {};
  const tap = transparency.taskbarTAP || {};
  const sm = transparency.startMenu || {};
  const ac = transparency.actionCenter || {};
  const aw = transparency.appWindows || {};
  const cm = transparency.contextMenus || {};

  return (
    <div>
      <h2 className="content__header">Transparency</h2>
      <p className="content__desc">Configure transparency effects for taskbar, Start Menu, Action Center, and app windows.</p>

      {/* Taskbar SWCA */}
      <div className="setting-group">
        <div className="setting-group__title">Taskbar (SWCA)</div>
        <div className="setting-row">
          <div className="setting-row__label">Enabled</div>
          <div className="setting-row__control">
            <Toggle value={tb.enabled} onChange={(v) => updateSub('taskbar', 'enabled', !!v)} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">Style</div>
          <div className="setting-row__control">
            <Select value={tb.style || 'clear'} onChange={(v) => updateSub('taskbar', 'style', v)}
              options={[
                { value: 'clear', label: 'Clear' }, { value: 'blur', label: 'Blur' },
                { value: 'acrylic', label: 'Acrylic' }, { value: 'opaque', label: 'Opaque' },
                { value: 'normal', label: 'Normal' }
              ]} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">Color Overlay</div>
          <div className="setting-row__control">
            <ColorPicker value={tb.color || '#00000000'} onChange={(v) => updateSub('taskbar', 'color', v)} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">All Monitors</div>
          <div className="setting-row__control">
            <Toggle value={tb.allMonitors} onChange={(v) => updateSub('taskbar', 'allMonitors', !!v)} />
          </div>
        </div>
        <button className="btn" onClick={() => apply('taskbar')} style={{ marginTop: 8 }}>Apply Taskbar</button>
      </div>

      {/* Taskbar TAP */}
      <div className="setting-group">
        <div className="setting-group__title">Taskbar TAP (XAML Injection)</div>
        <div className="setting-row">
          <div className="setting-row__label">Enabled</div>
          <div className="setting-row__control">
            <Toggle value={tap.enabled} onChange={(v) => updateSub('taskbarTAP', 'enabled', !!v)} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">Mode</div>
          <div className="setting-row__control">
            <Select value={tap.mode || 'Transparent'} onChange={(v) => updateSub('taskbarTAP', 'mode', v)}
              options={[
                { value: 'Transparent', label: 'Transparent' },
                { value: 'Acrylic', label: 'Acrylic' }
              ]} />
          </div>
        </div>
        <button className="btn" onClick={() => apply('taskbarTAP')} style={{ marginTop: 8 }}>Apply TAP</button>
      </div>

      {/* Start Menu */}
      <div className="setting-group">
        <div className="setting-group__title">Start Menu</div>
        <div className="setting-row">
          <div className="setting-row__label">Enabled</div>
          <div className="setting-row__control">
            <Toggle value={sm.enabled} onChange={(v) => updateSub('startMenu', 'enabled', !!v)} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">Mode</div>
          <div className="setting-row__control">
            <Select value={sm.mode || 'Transparent'} onChange={(v) => updateSub('startMenu', 'mode', v)}
              options={[
                { value: 'Transparent', label: 'Transparent' },
                { value: 'Acrylic', label: 'Acrylic' }
              ]} />
          </div>
        </div>
        <button className="btn" onClick={() => apply('startMenu')} style={{ marginTop: 8 }}>Apply Start Menu</button>
      </div>

      {/* Action Center */}
      <div className="setting-group">
        <div className="setting-group__title">Action Center</div>
        <div className="setting-row">
          <div className="setting-row__label">Enabled</div>
          <div className="setting-row__control">
            <Toggle value={ac.enabled} onChange={(v) => updateSub('actionCenter', 'enabled', !!v)} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">Mode</div>
          <div className="setting-row__control">
            <Select value={ac.mode || 'Transparent'} onChange={(v) => updateSub('actionCenter', 'mode', v)}
              options={[
                { value: 'Transparent', label: 'Transparent' },
                { value: 'Acrylic', label: 'Acrylic' }
              ]} />
          </div>
        </div>
        <button className="btn" onClick={() => apply('actionCenter')} style={{ marginTop: 8 }}>Apply Action Center</button>
      </div>

      {/* App Windows */}
      <div className="setting-group">
        <div className="setting-group__title">App Windows (BackdropWatcher)</div>
        <div className="setting-row">
          <div className="setting-row__label">Enabled</div>
          <div className="setting-row__control">
            <Toggle value={aw.enabled} onChange={(v) => updateSub('appWindows', 'enabled', !!v)} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">Backdrop</div>
          <div className="setting-row__control">
            <Select value={aw.backdrop || 'mica'} onChange={(v) => updateSub('appWindows', 'backdrop', v)}
              options={[
                { value: 'mica', label: 'Mica' },
                { value: 'acrylic', label: 'Acrylic' },
                { value: 'tabbed', label: 'Tabbed (Mica Alt)' }
              ]} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">Force Dark Mode</div>
          <div className="setting-row__control">
            <Toggle value={aw.darkMode} onChange={(v) => updateSub('appWindows', 'darkMode', !!v)} />
          </div>
        </div>
        <button className="btn" onClick={() => apply('appWindows')} style={{ marginTop: 8 }}>Apply Backdrops</button>
      </div>

      {/* Context Menus */}
      <div className="setting-group">
        <div className="setting-group__title">Context Menus</div>
        <div className="setting-row">
          <div>
            <div className="setting-row__label">Enable Backdrop</div>
            <div className="setting-row__desc">Requires App Windows to be enabled</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={cm.enabled} onChange={(v) => updateSub('contextMenus', 'enabled', !!v)} />
          </div>
        </div>
      </div>

      {/* Persistence */}
      <div className="setting-group">
        <div className="setting-group__title">Persistence</div>
        <div className="setting-row">
          <div>
            <div className="setting-row__label">Auto-restore on Login</div>
            <div className="setting-row__desc">Re-apply all transparency effects after login and process restart</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={transparency.persist} onChange={(v) => setField('transparency', 'persist', !!v)} />
          </div>
        </div>
      </div>
    </div>
  );
}
