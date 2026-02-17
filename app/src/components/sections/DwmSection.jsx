import useThemeStore from '../../store/themeStore';
import { useLivePreview } from '../../hooks/useLivePreview';
import AlphaColorPicker from '../controls/AlphaColorPicker';
import Slider from '../controls/Slider';
import Toggle from '../controls/Toggle';
import Select from '../controls/Select';

export default function DwmSection() {
  const dwm = useThemeStore((s) => s.theme.dwm);
  const setField = useThemeStore((s) => s.setField);
  useLivePreview('dwm');

  const update = (key, value) => setField('dwm', key, value);

  return (
    <div>
      <h2 className="content__header">Desktop Window Manager</h2>
      <p className="content__desc">Control DWM colorization, blur, glass effects, and window composition.</p>

      <div className="setting-group">
        <div className="setting-group__title">Colorization</div>

        <div className="setting-row">
          <div className="setting-row__label">Colorization Color</div>
          <div className="setting-row__control">
            <AlphaColorPicker value={dwm.colorizationColor} onChange={(v) => update('colorizationColor', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Afterglow Color</div>
          <div className="setting-row__control">
            <AlphaColorPicker value={dwm.colorizationAfterglow} onChange={(v) => update('colorizationAfterglow', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Inactive Window Accent</div>
          <div className="setting-row__control">
            <AlphaColorPicker value={dwm.accentColorInactive} onChange={(v) => update('accentColorInactive', v)} />
          </div>
        </div>
      </div>

      <div className="setting-group">
        <div className="setting-group__title">Balance & Intensity</div>

        <div className="setting-row">
          <div className="setting-row__label">Color Balance</div>
          <div className="setting-row__control">
            <Slider value={dwm.colorizationColorBalance} onChange={(v) => update('colorizationColorBalance', v)} suffix="%" />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Afterglow Balance</div>
          <div className="setting-row__control">
            <Slider value={dwm.colorizationAfterglowBalance} onChange={(v) => update('colorizationAfterglowBalance', v)} suffix="%" />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Blur Balance</div>
          <div className="setting-row__control">
            <Slider value={dwm.colorizationBlurBalance} onChange={(v) => update('colorizationBlurBalance', v)} suffix="%" />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Glass Reflection</div>
          <div className="setting-row__control">
            <Slider value={dwm.colorizationGlassReflectionIntensity} onChange={(v) => update('colorizationGlassReflectionIntensity', v)} suffix="%" />
          </div>
        </div>
      </div>

      <div className="setting-group">
        <div className="setting-group__title">Effects</div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">Aero Peek</div>
            <div className="setting-row__desc">Enable desktop preview on taskbar hover</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={dwm.enableAeroPeek} onChange={(v) => update('enableAeroPeek', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">Window Colorization</div>
            <div className="setting-row__desc">Enable color tinting on window frames</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={dwm.enableWindowColorization} onChange={(v) => update('enableWindowColorization', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">Force Effect Mode</div>
            <div className="setting-row__desc">Override system effect state</div>
          </div>
          <div className="setting-row__control">
            <Select value={dwm.forceEffectMode} onChange={(v) => update('forceEffectMode', v)}
              options={[
                { value: 0, label: 'Default' },
                { value: 1, label: 'Force Disable' },
                { value: 2, label: 'Force Enable' }
              ]} />
          </div>
        </div>
      </div>
    </div>
  );
}
