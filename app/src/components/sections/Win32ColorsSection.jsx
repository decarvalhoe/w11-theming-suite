import useThemeStore from '../../store/themeStore';
import { useLivePreview } from '../../hooks/useLivePreview';
import ColorPicker from '../controls/ColorPicker';

const COLOR_GROUPS = [
  {
    title: 'Window',
    items: [
      { key: 'Window', label: 'Background' }, { key: 'WindowText', label: 'Text' },
      { key: 'WindowFrame', label: 'Frame' }, { key: 'Scrollbar', label: 'Scrollbar' }
    ]
  },
  {
    title: 'Title Bar',
    items: [
      { key: 'ActiveTitle', label: 'Active Title' }, { key: 'InactiveTitle', label: 'Inactive Title' },
      { key: 'TitleText', label: 'Title Text' }, { key: 'InactiveTitleText', label: 'Inactive Title Text' },
      { key: 'GradientActiveTitle', label: 'Gradient Active' }, { key: 'GradientInactiveTitle', label: 'Gradient Inactive' }
    ]
  },
  {
    title: 'Buttons',
    items: [
      { key: 'ButtonFace', label: 'Face' }, { key: 'ButtonText', label: 'Text' },
      { key: 'ButtonShadow', label: 'Shadow' }, { key: 'ButtonHilight', label: 'Highlight' },
      { key: 'ButtonDkShadow', label: 'Dark Shadow' }, { key: 'ButtonLight', label: 'Light' }
    ]
  },
  {
    title: 'Menu',
    items: [
      { key: 'Menu', label: 'Background' }, { key: 'MenuText', label: 'Text' },
      { key: 'MenuHilight', label: 'Highlight' }, { key: 'MenuBar', label: 'Menu Bar' }
    ]
  },
  {
    title: 'Selection & Links',
    items: [
      { key: 'Hilight', label: 'Selection BG' }, { key: 'HilightText', label: 'Selection Text' },
      { key: 'HotTrackingColor', label: 'Links / Hot Track' }, { key: 'GrayText', label: 'Disabled Text' }
    ]
  },
  {
    title: 'Misc',
    items: [
      { key: 'Background', label: 'Desktop BG' }, { key: 'AppWorkspace', label: 'App Workspace' },
      { key: 'ActiveBorder', label: 'Active Border' }, { key: 'InactiveBorder', label: 'Inactive Border' },
      { key: 'InfoText', label: 'Tooltip Text' }, { key: 'InfoWindow', label: 'Tooltip BG' }
    ]
  }
];

export default function Win32ColorsSection() {
  const colors = useThemeStore((s) => s.theme.win32Colors);
  const setField = useThemeStore((s) => s.setField);
  useLivePreview('win32Colors', 200);

  const update = (key, value) => setField('win32Colors', key, value);

  return (
    <div>
      <h2 className="content__header">Win32 Colors</h2>
      <p className="content__desc">
        Classic Windows color scheme (Control Panel\Colors). Each value is in "R G B" format.
        These affect legacy Win32 applications, dialogs, and some shell surfaces.
      </p>

      {COLOR_GROUPS.map((group) => (
        <div key={group.title} className="setting-group">
          <div className="setting-group__title">{group.title}</div>
          <div className="color-grid">
            {group.items.map((item) => (
              <div key={item.key} className="color-grid__item">
                <ColorPicker value={colors?.[item.key] || '0 0 0'} format="rgb"
                  onChange={(v) => update(item.key, v)} />
                <span className="color-grid__label">{item.label}</span>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
