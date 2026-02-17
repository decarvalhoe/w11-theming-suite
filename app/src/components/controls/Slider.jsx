/**
 * Slider — Numeric range input with label showing the current value.
 */
export default function Slider({ value, onChange, min = 0, max = 100, step = 1, suffix = '' }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 10, minWidth: 180 }}>
      <input
        type="range"
        min={min}
        max={max}
        step={step}
        value={value ?? min}
        onChange={(e) => onChange(Number(e.target.value))}
        style={{ flex: 1, accentColor: 'var(--accent)', cursor: 'pointer' }}
      />
      <span style={{
        fontSize: 12, color: 'var(--text-secondary)',
        minWidth: 36, textAlign: 'right', fontFamily: 'var(--font-mono)'
      }}>
        {value ?? min}{suffix}
      </span>
    </div>
  );
}
