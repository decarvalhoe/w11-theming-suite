import ColorPicker from './ColorPicker';

/**
 * PaletteEditor — Edit the 8-slot AccentPalette array.
 */
const SLOT_LABELS = [
  'Lightest', 'Light', 'Medium Light', 'Normal',
  'Medium Dark', 'Dark', 'Darkest', 'Highlight'
];

export default function PaletteEditor({ value = [], onChange }) {
  const palette = [...(value || [])];
  while (palette.length < 8) palette.push('#000000');

  const handleSlotChange = (index, color) => {
    const next = [...palette];
    next[index] = color;
    onChange(next);
  };

  return (
    <div>
      <div className="palette-row" style={{ flexWrap: 'wrap', gap: 10 }}>
        {palette.map((color, i) => (
          <div key={i} style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
            <ColorPicker value={color} onChange={(c) => handleSlotChange(i, c)} />
            <span style={{ fontSize: 10, color: 'var(--text-muted)' }}>{SLOT_LABELS[i]}</span>
          </div>
        ))}
      </div>
    </div>
  );
}
