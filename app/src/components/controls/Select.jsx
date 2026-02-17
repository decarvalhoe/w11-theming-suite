/**
 * Select — Dropdown for enum values.
 * options: [{ value: 0, label: 'Left' }, { value: 1, label: 'Center' }]
 */
export default function Select({ value, onChange, options = [] }) {
  return (
    <select
      value={value ?? ''}
      onChange={(e) => {
        const opt = options.find((o) => String(o.value) === e.target.value);
        onChange(opt ? opt.value : e.target.value);
      }}
      style={{
        background: 'var(--bg-tertiary)', color: 'var(--text-primary)',
        border: '1px solid var(--border-color)', borderRadius: 'var(--radius-sm)',
        padding: '5px 10px', fontSize: 12, cursor: 'pointer', minWidth: 120
      }}
    >
      {options.map((opt) => (
        <option key={String(opt.value)} value={String(opt.value)}>{opt.label}</option>
      ))}
    </select>
  );
}
