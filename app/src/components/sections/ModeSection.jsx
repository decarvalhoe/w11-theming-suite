import useThemeStore from '../../store/themeStore';
import { useLivePreview } from '../../hooks/useLivePreview';
import Toggle from '../controls/Toggle';

export default function ModeSection() {
  const mode = useThemeStore((s) => s.theme.mode);
  const setField = useThemeStore((s) => s.setField);
  useLivePreview('mode');

  return (
    <div>
      <h2 className="content__header">Mode</h2>
      <p className="content__desc">Toggle between dark and light themes for apps and the system shell.</p>

      <div className="setting-group">
        <div className="setting-group__title">Theme Mode</div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">Apps Dark Mode</div>
            <div className="setting-row__desc">Use dark theme for UWP and WinUI applications</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={mode.appsUseLightTheme === 0 ? 1 : 0}
              onChange={(v) => setField('mode', 'appsUseLightTheme', v ? 0 : 1)} />
          </div>
        </div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">System Dark Mode</div>
            <div className="setting-row__desc">Use dark theme for taskbar, Start Menu, and system shell</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={mode.systemUsesLightTheme === 0 ? 1 : 0}
              onChange={(v) => setField('mode', 'systemUsesLightTheme', v ? 0 : 1)} />
          </div>
        </div>
      </div>
    </div>
  );
}
