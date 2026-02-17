import useThemeStore from '../../store/themeStore';
import FilePicker from '../controls/FilePicker';

const CURSOR_ROLES = [
  { key: 'Arrow', label: 'Normal Select' }, { key: 'Help', label: 'Help Select' },
  { key: 'AppStarting', label: 'Working in Background' }, { key: 'Wait', label: 'Busy' },
  { key: 'Crosshair', label: 'Precision Select' }, { key: 'IBeam', label: 'Text Select' },
  { key: 'NWPen', label: 'Handwriting' }, { key: 'No', label: 'Unavailable' },
  { key: 'SizeNS', label: 'Vertical Resize' }, { key: 'SizeWE', label: 'Horizontal Resize' },
  { key: 'SizeNWSE', label: 'Diagonal Resize (NW-SE)' }, { key: 'SizeNESW', label: 'Diagonal Resize (NE-SW)' },
  { key: 'SizeAll', label: 'Move' }, { key: 'UpArrow', label: 'Alternate Select' },
  { key: 'Hand', label: 'Link Select' }
];

const CURSOR_FILTERS = [{ name: 'Cursors', extensions: ['cur', 'ani'] }];

export default function CursorsSection() {
  const cursors = useThemeStore((s) => s.theme.cursors);
  const setField = useThemeStore((s) => s.setField);
  const setLastAction = useThemeStore((s) => s.setLastAction);

  const updateRole = (role, path) => {
    const roles = { ...(cursors.roles || {}), [role]: path };
    setField('cursors', 'roles', roles);
  };

  const handleApply = async () => {
    if (!window.themeAPI) return;
    try {
      await window.themeAPI.applyCursors(cursors);
      setLastAction('Applied cursor scheme');
    } catch (err) {
      setLastAction('Cursor error: ' + err.message, false);
    }
  };

  return (
    <div>
      <h2 className="content__header">Cursors</h2>
      <p className="content__desc">Configure the 15 cursor roles. Select .cur or .ani files for each.</p>

      <div className="setting-group">
        <div className="setting-group__title">Scheme Name</div>
        <div className="setting-row">
          <div className="setting-row__label">Name</div>
          <div className="setting-row__control">
            <input type="text" value={cursors.schemeName || ''}
              onChange={(e) => setField('cursors', 'schemeName', e.target.value)}
              placeholder="My Cursor Scheme"
              style={{
                background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
                borderRadius: 'var(--radius-sm)', padding: '5px 10px', color: 'var(--text-primary)', fontSize: 12
              }} />
          </div>
        </div>
      </div>

      <div className="setting-group">
        <div className="setting-group__title">Cursor Roles (15)</div>
        {CURSOR_ROLES.map((role) => (
          <div key={role.key} className="setting-row">
            <div>
              <div className="setting-row__label">{role.label}</div>
              <div className="setting-row__desc">{role.key}</div>
            </div>
            <div className="setting-row__control">
              <FilePicker value={cursors.roles?.[role.key] || ''} filters={CURSOR_FILTERS}
                onChange={(v) => updateRole(role.key, v)} />
            </div>
          </div>
        ))}
      </div>

      <button className="btn btn--primary" onClick={handleApply} style={{ marginTop: 8 }}>
        Apply Cursor Scheme
      </button>
    </div>
  );
}
