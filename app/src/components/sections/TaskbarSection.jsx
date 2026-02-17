import useThemeStore from '../../store/themeStore';
import { useLivePreview } from '../../hooks/useLivePreview';
import Toggle from '../controls/Toggle';
import Select from '../controls/Select';

export default function TaskbarSection() {
  const taskbar = useThemeStore((s) => s.theme.taskbar);
  const setField = useThemeStore((s) => s.setField);
  useLivePreview('taskbar');

  const update = (key, value) => setField('taskbar', key, value);

  return (
    <div>
      <h2 className="content__header">Taskbar</h2>
      <p className="content__desc">Configure taskbar alignment, visibility, and search mode.</p>

      <div className="setting-group">
        <div className="setting-group__title">Layout</div>

        <div className="setting-row">
          <div className="setting-row__label">Icon Alignment</div>
          <div className="setting-row__control">
            <Select value={taskbar.alignment} onChange={(v) => update('alignment', v)}
              options={[
                { value: 0, label: 'Left' },
                { value: 1, label: 'Center' }
              ]} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Search Mode</div>
          <div className="setting-row__control">
            <Select value={taskbar.showSearch} onChange={(v) => update('showSearch', v)}
              options={[
                { value: 0, label: 'Hidden' },
                { value: 1, label: 'Icon Only' },
                { value: 2, label: 'Icon + Label' },
                { value: 3, label: 'Search Box' }
              ]} />
          </div>
        </div>
      </div>

      <div className="setting-group">
        <div className="setting-group__title">Buttons</div>

        <div className="setting-row">
          <div className="setting-row__label">Show Task View</div>
          <div className="setting-row__control">
            <Toggle value={taskbar.showTaskView} onChange={(v) => update('showTaskView', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Show Widgets</div>
          <div className="setting-row__control">
            <Toggle value={taskbar.showWidgets} onChange={(v) => update('showWidgets', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div className="setting-row__label">Show Chat</div>
          <div className="setting-row__control">
            <Toggle value={taskbar.showChat} onChange={(v) => update('showChat', v)} />
          </div>
        </div>

        <div className="setting-row">
          <div>
            <div className="setting-row__label">OLED Transparency</div>
            <div className="setting-row__desc">Enhanced transparency optimized for OLED displays</div>
          </div>
          <div className="setting-row__control">
            <Toggle value={taskbar.useOLEDTransparency} onChange={(v) => update('useOLEDTransparency', v)} />
          </div>
        </div>
      </div>
    </div>
  );
}
