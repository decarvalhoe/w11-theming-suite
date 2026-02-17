import useThemeStore from '../../store/themeStore';
import FilePicker from '../controls/FilePicker';

const SYSTEM_EVENTS = [
  'SystemAsterisk', 'SystemExclamation', 'SystemHand', 'SystemNotification', 'SystemQuestion',
  'SystemExit', 'WindowsLogon', 'WindowsLogoff', 'WindowsUAC',
  'DeviceConnect', 'DeviceDisconnect', 'DeviceFail',
  'MailBeep', 'Close', 'Open', 'Maximize', 'Minimize', 'RestoreUp', 'RestoreDown',
  'MenuCommand', 'MenuPopup', 'PrintComplete', 'CriticalBatteryAlarm', 'LowBatteryAlarm'
];

const EXPLORER_EVENTS = ['BlockedPopup', 'EmptyRecycleBin', 'FeedDiscovered', 'Navigating', 'SecurityBand'];

const WAV_FILTERS = [{ name: 'WAV Audio', extensions: ['wav'] }];

export default function SoundsSection() {
  const sounds = useThemeStore((s) => s.theme.sounds);
  const setField = useThemeStore((s) => s.setField);
  const setLastAction = useThemeStore((s) => s.setLastAction);

  const updateEvent = (event, path) => {
    const events = { ...(sounds.events || {}), [event]: path };
    setField('sounds', 'events', events);
  };

  const updateExplorerEvent = (event, path) => {
    const events = { ...(sounds.explorerEvents || {}), [event]: path };
    setField('sounds', 'explorerEvents', events);
  };

  const handleApply = async () => {
    if (!window.themeAPI) return;
    try {
      await window.themeAPI.applySounds(sounds);
      setLastAction('Applied sound scheme');
    } catch (err) {
      setLastAction('Sound error: ' + err.message, false);
    }
  };

  return (
    <div>
      <h2 className="content__header">Sounds</h2>
      <p className="content__desc">Configure system and Explorer sound events with .wav files.</p>

      <div className="setting-group">
        <div className="setting-group__title">Scheme Name</div>
        <div className="setting-row">
          <div className="setting-row__label">Name</div>
          <div className="setting-row__control">
            <input type="text" value={sounds.schemeName || ''}
              onChange={(e) => setField('sounds', 'schemeName', e.target.value)}
              placeholder="My Sound Scheme"
              style={{
                background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
                borderRadius: 'var(--radius-sm)', padding: '5px 10px', color: 'var(--text-primary)', fontSize: 12
              }} />
          </div>
        </div>
      </div>

      <div className="setting-group">
        <div className="setting-group__title">System Events ({SYSTEM_EVENTS.length})</div>
        {SYSTEM_EVENTS.map((event) => (
          <div key={event} className="setting-row">
            <div className="setting-row__label">{event}</div>
            <div className="setting-row__control">
              <FilePicker value={sounds.events?.[event] || ''} filters={WAV_FILTERS}
                onChange={(v) => updateEvent(event, v)} />
            </div>
          </div>
        ))}
      </div>

      <div className="setting-group">
        <div className="setting-group__title">Explorer Events ({EXPLORER_EVENTS.length})</div>
        {EXPLORER_EVENTS.map((event) => (
          <div key={event} className="setting-row">
            <div className="setting-row__label">{event}</div>
            <div className="setting-row__control">
              <FilePicker value={sounds.explorerEvents?.[event] || ''} filters={WAV_FILTERS}
                onChange={(v) => updateExplorerEvent(event, v)} />
            </div>
          </div>
        ))}
      </div>

      <button className="btn btn--primary" onClick={handleApply} style={{ marginTop: 8 }}>
        Apply Sound Scheme
      </button>
    </div>
  );
}
