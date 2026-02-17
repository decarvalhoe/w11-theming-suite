import useThemeStore from '../../store/themeStore';
import { useLivePreview } from '../../hooks/useLivePreview';
import ColorPicker from '../controls/ColorPicker';
import Toggle from '../controls/Toggle';
import PaletteEditor from '../controls/PaletteEditor';

export default function AccentColorSection() {
  const accent = useThemeStore((s) => s.theme.accentColor);
  const setField = useThemeStore((s) => s.setField);
  useLivePreview('accentColor');

  const update = (key, value) => setField('accentColor', key, value);

  return (
    <div>
      <h2 className="content__header">Accent Color</h2>
      <p className="content__desc">Configure the system accent color, prevalence, and the 8-slot palette.</p>

      <div className="setting-group">
        <div className="setting-group__title">Primary Accent</div>

        <div className="setting-row">
          <div className="setting-row__label">Accent Color</div>
          <div className="setting-row__control">
            <ColorPicker value={accent.color} onChange={(v) => update('color', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">Show on Taskbar & Start</div>
            <div className="setting-row__desc">Apply accent color to taskbar and Start Menu surfaces</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={accent.colorPrevalence} onChange={(v) => update('colorPrevalence', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">Show on Title Bars</div>
            <div className="setting-row__desc">Apply accent color to window title bars and borders</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={accent.colorPrevalenceTitleBars} onChange={(v) => update('colorPrevalenceTitleBars', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">Enable Transparency</div>
            <div className="setting-row__desc">Enable transparency effects in the shell</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={accent.enableTransparency} onChange={(v) => update('enableTransparency', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">Auto from Wallpaper</div>
            <div className="setting-row__desc">Automatically pick accent color from the desktop wallpaper</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={accent.autoColorization} onChange={(v) => update('autoColorization', v)} />
          </div>
        </div>
      </div>

      <div className="setting-group">
        <div className="setting-group__title">Accent Palette (8 Slots)</div>
        <p style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 12 }}>
          These 8 color slots control gradient shading across system surfaces. Slots 0-2 affect light variants (toggle icons),
          slots 3-6 affect dark variants (taskbar, Start), slot 7 is the highlight color.
        </p>
        <PaletteEditor value={accent.accentPalette} onChange={(v) => update('accentPalette', v)} />
      </div>
    </div>
  );
}
