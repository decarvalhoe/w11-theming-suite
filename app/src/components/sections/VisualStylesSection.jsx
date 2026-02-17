import useThemeStore from '../../store/themeStore';
import FilePicker from '../controls/FilePicker';
import AlphaColorPicker from '../controls/AlphaColorPicker';
import Toggle from '../controls/Toggle';

export default function VisualStylesSection() {
  const vs = useThemeStore((s) => s.theme.visualStyles);
  const setField = useThemeStore((s) => s.setField);

  const update = (key, value) => setField('visualStyles', key, value);

  return (
    <div>
      <h2 className="content__header">Visual Styles</h2>
      <p className="content__desc">Configure the .msstyles visual style, colorization, and composition settings.</p>

      <div className="setting-group">
        <div className="setting-group__title">Style File</div>

        <div className="setting-row">
          <div className="setting-row__label">.msstyles Path</div>
          <div className="setting-row__control">
            <FilePicker value={vs.path} onChange={(v) => update('path', v)}
              filters={[{ name: 'Visual Styles', extensions: ['msstyles'] }]} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Color Style</div>
          <div className="setting-row__control">
            <input type="text" value={vs.colorStyle || ''} onChange={(e) => update('colorStyle', e.target.value)}
              style={{
                background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
                borderRadius: 'var(--radius-sm)', padding: '5px 10px', color: 'var(--text-primary)', fontSize: 12, width: 140
              }} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Size</div>
          <div className="setting-row__control">
            <input type="text" value={vs.size || ''} onChange={(e) => update('size', e.target.value)}
              style={{
                background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
                borderRadius: 'var(--radius-sm)', padding: '5px 10px', color: 'var(--text-primary)', fontSize: 12, width: 140
              }} />
          </div>
        </div>
      </div>

      <div className="setting-group">
        <div className="setting-group__title">Composition</div>

        <div className="setting-row">
          <div className="setting-row__label">Colorization Color</div>
          <div className="setting-row__control">
            <AlphaColorPicker value={vs.colorizationColor} onChange={(v) => update('colorizationColor', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Enable Transparency</div>
          <div className="setting-row__control">
            <Toggle value={vs.transparency} onChange={(v) => update('transparency', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Enable Composition</div>
          <div className="setting-row__control">
            <Toggle value={vs.composition} onChange={(v) => update('composition', v)} />
          </div>
        </div>
      </div>
    </div>
  );
}
