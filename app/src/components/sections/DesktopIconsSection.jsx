import useThemeStore from '../../store/themeStore';
import FilePicker from '../controls/FilePicker';

const ICONS = [
  { key: 'computer', label: 'This PC' },
  { key: 'documents', label: 'Documents' },
  { key: 'network', label: 'Network' },
  { key: 'recycleBinFull', label: 'Recycle Bin (Full)' },
  { key: 'recycleBinEmpty', label: 'Recycle Bin (Empty)' }
];

const ICO_FILTERS = [{ name: 'Icons', extensions: ['ico'] }];

export default function DesktopIconsSection() {
  const icons = useThemeStore((s) => s.theme.desktopIcons);
  const setField = useThemeStore((s) => s.setField);

  const update = (key, value) => setField('desktopIcons', key, value);

  return (
    <div>
      <h2 className="content__header">Desktop Icons</h2>
      <p className="content__desc">Assign custom .ico files for the standard desktop icons.</p>

      <div className="setting-group">
        <div className="setting-group__title">Icon Paths</div>
        {ICONS.map((icon) => (
          <div key={icon.key} className="setting-row">
            <div className="setting-row__label">{icon.label}</div>
            <div className="setting-row__control">
              <FilePicker value={icons?.[icon.key] || ''} filters={ICO_FILTERS}
                onChange={(v) => update(icon.key, v)} />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
