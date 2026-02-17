/**
 * Toggle — Switch for boolean (true/false) or integer (0/1) values.
 */
export default function Toggle({ value, onChange, disabled = false }) {
  const isOn = value === 1 || value === true;
  return (
    <button
      className={`toggle ${isOn ? 'toggle--on' : ''} ${disabled ? 'toggle--disabled' : ''}`}
      onClick={() => !disabled && onChange(isOn ? 0 : 1)}
      disabled={disabled}
      type="button"
      style={{
        position: 'relative', width: 40, height: 22, borderRadius: 11, border: 'none',
        background: isOn ? 'var(--accent)' : '#444', cursor: disabled ? 'not-allowed' : 'pointer',
        transition: 'background 0.15s', flexShrink: 0
      }}
    >
      <span style={{
        position: 'absolute', top: 2, left: isOn ? 20 : 2,
        width: 18, height: 18, borderRadius: '50%', background: '#fff',
        transition: 'left 0.15s', boxShadow: '0 1px 3px rgba(0,0,0,0.4)'
      }} />
    </button>
  );
}
