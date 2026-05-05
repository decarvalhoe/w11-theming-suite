import { useState, useRef, useEffect, useCallback } from 'react';
import { HexColorPicker } from 'react-colorful';

/**
 * AlphaColorPicker — For "0xAARRGGBB" DWM color values.
 * Displays a color swatch + alpha slider + editable hex text field.
 */
export default function AlphaColorPicker({ value, onChange }) {
  const [open, setOpen] = useState(false);
  const [textInput, setTextInput] = useState('');
  const popover = useRef(null);

  // Parse "0xAARRGGBB" → { alpha, hex }
  const parse = useCallback((val) => {
    if (!val || typeof val !== 'string') return { alpha: 255, hex: '#000000' };
    const clean = val.replace(/^0x/i, '').padStart(8, '0');
    const a = parseInt(clean.substring(0, 2), 16);
    const r = clean.substring(2, 4);
    const g = clean.substring(4, 6);
    const b = clean.substring(6, 8);
    return { alpha: isNaN(a) ? 255 : a, hex: `#${r}${g}${b}` };
  }, []);

  // Format back to "0xAARRGGBB"
  const format = useCallback((hex, alpha) => {
    const h = hex.replace('#', '');
    const a = Math.max(0, Math.min(255, alpha)).toString(16).padStart(2, '0').toUpperCase();
    return `0x${a}${h.toUpperCase()}`;
  }, []);

  const { alpha, hex } = parse(value);

  // Sync text input when external value changes
  useEffect(() => {
    setTextInput(value || '0x00000000');
  }, [value]);

  const handleColorChange = useCallback((newHex) => {
    onChange(format(newHex, alpha));
  }, [alpha, onChange, format]);

  const handleAlphaChange = useCallback((e) => {
    onChange(format(hex, Number(e.target.value)));
  }, [hex, onChange, format]);

  // Handle text field commit
  const commitText = useCallback(() => {
    const trimmed = textInput.trim();
    // Accept 0xAARRGGBB or AARRGGBB (8 hex chars with optional 0x prefix)
    if (/^(0x)?[0-9a-fA-F]{8}$/i.test(trimmed)) {
      const clean = trimmed.replace(/^0x/i, '');
      onChange('0x' + clean.toUpperCase());
      return;
    }
    // Accept #RRGGBB (6 hex chars) — use current alpha
    if (/^#?[0-9a-fA-F]{6}$/.test(trimmed)) {
      const hexPart = trimmed.replace('#', '');
      const a = Math.max(0, Math.min(255, alpha)).toString(16).padStart(2, '0').toUpperCase();
      onChange(`0x${a}${hexPart.toUpperCase()}`);
      return;
    }
    // Invalid — revert
    setTextInput(value || '0x00000000');
  }, [textInput, alpha, value, onChange]);

  useEffect(() => {
    if (!open) return;
    const handleClick = (e) => {
      if (popover.current && !popover.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [open]);

  return (
    <div style={{ position: 'relative', display: 'inline-flex', alignItems: 'center', gap: 8 }}>
      <button type="button" onClick={() => setOpen(!open)} style={{
        width: 28, height: 28, borderRadius: 'var(--radius-sm)',
        border: '1px solid var(--border-color)', background: hex,
        opacity: Math.max(0.15, alpha / 255), cursor: 'pointer', flexShrink: 0
      }} />
      <input
        type="text"
        value={textInput}
        onChange={(e) => setTextInput(e.target.value)}
        onBlur={commitText}
        onKeyDown={(e) => { if (e.key === 'Enter') { e.target.blur(); } }}
        spellCheck={false}
        style={{
          fontSize: 11, fontFamily: 'var(--font-mono)', color: 'var(--text-primary)',
          background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
          borderRadius: 'var(--radius-sm)', padding: '3px 6px', width: 90, outline: 'none'
        }}
      />
      {open && (
        <div ref={popover} style={{
          position: 'absolute', top: 34, left: 0, zIndex: 1000,
          background: 'var(--bg-secondary)', border: '1px solid var(--border-color)',
          borderRadius: 'var(--radius)', padding: 10, boxShadow: '0 8px 24px rgba(0,0,0,0.5)'
        }}>
          <HexColorPicker color={hex} onChange={handleColorChange} />
          <div style={{ marginTop: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
            <label style={{ fontSize: 11, color: 'var(--text-muted)' }}>Alpha</label>
            <input type="range" min={0} max={255} value={alpha} onChange={handleAlphaChange}
              style={{ flex: 1, accentColor: 'var(--accent)' }} />
            <span style={{ fontSize: 11, fontFamily: 'var(--font-mono)', color: 'var(--text-secondary)', minWidth: 24 }}>
              {alpha}
            </span>
          </div>
        </div>
      )}
    </div>
  );
}
