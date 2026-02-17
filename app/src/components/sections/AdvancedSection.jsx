import useThemeStore from '../../store/themeStore';
import RegistryOverrideEditor from '../controls/RegistryOverrideEditor';

export default function AdvancedSection() {
  const advanced = useThemeStore((s) => s.theme.advanced);
  const setField = useThemeStore((s) => s.setField);

  return (
    <div>
      <h2 className="content__header">Advanced</h2>
      <p className="content__desc">Add raw registry overrides and configure third-party tool paths.</p>

      <div className="setting-group">
        <div className="setting-group__title">Registry Overrides</div>
        <p style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 12 }}>
          Add arbitrary registry values that will be applied with the theme.
          These run after all standard sections.
        </p>
        <RegistryOverrideEditor
          value={advanced?.registryOverrides || []}
          onChange={(v) => setField('advanced', 'registryOverrides', v)} />
      </div>

      <div className="setting-group">
        <div className="setting-group__title">Third-Party Tools</div>
        <div className="setting-row">
          <div className="setting-row__label">msstyle Editor Path</div>
          <div className="setting-row__control">
            <input type="text" value={advanced?.thirdParty?.msstyleEditorPath || ''}
              onChange={(e) => setField('advanced', 'thirdParty', {
                ...(advanced?.thirdParty || {}), msstyleEditorPath: e.target.value
              })}
              placeholder="C:\Tools\msstyleEditor.exe"
              style={{
                background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
                borderRadius: 'var(--radius-sm)', padding: '5px 10px', color: 'var(--text-primary)',
                fontSize: 12, fontFamily: 'var(--font-mono)', width: 280
              }} />
          </div>
        </div>
        <div className="setting-row">
          <div className="setting-row__label">Mica For Everyone Path</div>
          <div className="setting-row__control">
            <input type="text" value={advanced?.thirdParty?.micaForEveryonePath || ''}
              onChange={(e) => setField('advanced', 'thirdParty', {
                ...(advanced?.thirdParty || {}), micaForEveryonePath: e.target.value
              })}
              placeholder="C:\Tools\MicaForEveryone.exe"
              style={{
                background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
                borderRadius: 'var(--radius-sm)', padding: '5px 10px', color: 'var(--text-primary)',
                fontSize: 12, fontFamily: 'var(--font-mono)', width: 280
              }} />
          </div>
        </div>
      </div>
    </div>
  );
}
