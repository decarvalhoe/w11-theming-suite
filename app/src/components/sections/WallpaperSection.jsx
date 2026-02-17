import useThemeStore from '../../store/themeStore';
import Select from '../controls/Select';
import FilePicker from '../controls/FilePicker';
import ColorPicker from '../controls/ColorPicker';

export default function WallpaperSection() {
  const wallpaper = useThemeStore((s) => s.theme.wallpaper);
  const setField = useThemeStore((s) => s.setField);
  const setLastAction = useThemeStore((s) => s.setLastAction);

  const update = (key, value) => setField('wallpaper', key, value);

  const handleApply = async () => {
    if (!window.themeAPI) return;
    try {
      await window.themeAPI.applyWallpaper(wallpaper);
      setLastAction('Applied wallpaper');
    } catch (err) {
      setLastAction('Wallpaper error: ' + err.message, false);
    }
  };

  return (
    <div>
      <h2 className="content__header">Wallpaper</h2>
      <p className="content__desc">Configure the desktop wallpaper mode, image, and fit style.</p>

      <div className="setting-group">
        <div className="setting-group__title">Wallpaper Settings</div>

        <div className="setting-row">
          <div className="setting-row__label">Mode</div>
          <div className="setting-row__control">
            <Select value={wallpaper.mode} onChange={(v) => update('mode', v)}
              options={[
                { value: 'static', label: 'Static Image' },
                { value: 'slideshow', label: 'Slideshow' },
                { value: 'solid', label: 'Solid Color' }
              ]} />
          </div>
        </div>

        {wallpaper.mode !== 'solid' && (
          <div className="setting-row">
            <div className="setting-row__label">Image Path</div>
            <div className="setting-row__control">
              <FilePicker value={wallpaper.path} onChange={(v) => update('path', v)}
                filters={[{ name: 'Images', extensions: ['jpg', 'jpeg', 'png', 'bmp', 'webp'] }]} />
            </div>
          </div>
        )}

        <div className="setting-row">
          <div className="setting-row__label">Fit Style</div>
          <div className="setting-row__control">
            <Select value={wallpaper.style} onChange={(v) => update('style', v)}
              options={[
                { value: 10, label: 'Fill' },
                { value: 6, label: 'Fit' },
                { value: 2, label: 'Stretch' },
                { value: 0, label: 'Center' }
              ]} />
          </div>
        </div>

        {wallpaper.mode === 'solid' && (
          <div className="setting-row">
            <div className="setting-row__label">Solid Color</div>
            <div className="setting-row__control">
              <ColorPicker value={wallpaper.color || '#000000'} onChange={(v) => update('color', v)} />
            </div>
          </div>
        )}
      </div>

      <button className="btn btn--primary" onClick={handleApply} style={{ marginTop: 8 }}>
        Apply Wallpaper
      </button>
    </div>
  );
}
