import { useState, useRef, useEffect, useCallback } from 'react';
import { HexColorPicker } from 'react-colorful';

/**
 * ColorPicker — Supports #RRGGBB hex and "R G B" (Win32) formats.
 * Includes an editable hex text field.
 *
 * @param {string} value     – Current value (hex "#RRGGBB" or "R G B")
 * @param {function} onChange – Callback with the new value in the same format
 * @param {'hex'|'rgb'} format – Output format (default: 'hex')
 */
export default function ColorPicker({ value, onChange, format = 'hex' }) {
  const [open, setOpen] = useState(false);
  const [textInput, setTextInput] = useState('');
  const popover = useRef(null);

  // Convert "R G B" → "#RRGGBB"
  const toHex = useCallback((val) => {
    if (!val) return '#000000';
    if (val.startsWith('#')) return val.length >= 7 ? val.substring(0, 7) : val;
    const parts = val.split(' ').map(Number);
    if (parts.length === 3 && parts.every((n) => !isNaN(n))) {
      return '#' + parts.map((n) => Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0')).join('');
    }
    return '#000000';
  }, []);

  // Convert "#RRGGBB" → "R G B"
  const toRgbString = useCallback((hex) => {
    const h = hex.replace('#', '');
    const r = parseInt(h.substring(0, 2), 16);
    const g = parseInt(h.substring(2, 4), 16);
    const b = parseInt(h.substring(4, 6), 16);
    return `${r} ${g} ${b}`;
  }, []);

  const hexValue = toHex(value);

  // Sync text input when external value changes
  useEffect(() => {
    setTextInput(format === 'rgb' ? (value || '0 0 0') : hexValue);
  }, [value, hexValue, format]);

  const handleChange = useCallback((hex) => {
    if (format === 'rgb') {
      onChange(toRgbString(hex));
    } else {
      onChange(hex);
    }
  }, [format, onChange, toRgbString]);

  // Handle hex text field commit
  const commitText = useCallback(() => {
    const trimmed = textInput.trim();
    if (format === 'rgb') {
      // Accept "R G B" format
      const parts = trimmed.split(/[\s,]+/).map(Number);
      if (parts.length === 3 && parts.every((n) => !isNaN(n) && n >= 0 && n <= 255)) {
        onChange(`${parts[0]} ${parts[1]} ${parts[2]}`);
        return;
      }
      // Also accept hex → convert to RGB
      if (/^#?[0-9a-fA-F]{6}$/.test(trimmed)) {
        const hex = trimmed.startsWith('#') ? trimmed : '#' + trimmed;
        onChange(toRgbString(hex));
        return;
      }
    } else {
      // Accept #RRGGBB or RRGGBB
      if (/^#?[0-9a-fA-F]{6}$/.test(trimmed)) {
        const hex = trimmed.startsWith('#') ? trimmed : '#' + trimmed;
        onChange(hex.toLowerCase());
        return;
      }
    }
    // Invalid input — revert
    setTextInput(format === 'rgb' ? (value || '0 0 0') : hexValue);
  }, [textInput, format, onChange, toRgbString, value, hexValue]);

  // Close popover on outside click
  useEffect(() => {
    if (!open) return;
    const handleClick = (e) => {
      if (popover.current && !popover.current.contains(e.target)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [open]);

  return (
    <div style={{ position: 'relative', display: 'inline-flex', alignItems: 'center', gap: 8 }}>
      <button
        type="button"
        onClick={() => setOpen(!open)}
        style={{
          width: 28, height: 28, borderRadius: 'var(--radius-sm)',
          border: '1px solid var(--border-color)', background: hexValue,
          cursor: 'pointer', flexShrink: 0
        }}
      />
      <input
        type="text"
        value={textInput}
        onChange={(e) => setTextInput(e.target.value)}
        onBlur={commitText}
        onKeyDown={(e) => { if (e.key === 'Enter') { e.target.blur(); } }}
        spellCheck={false}
        style={{
          fontSize: 12, fontFamily: 'var(--font-mono)', color: 'var(--text-primary)',
          background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
          borderRadius: 'var(--radius-sm)', padding: '3px 6px', minWidth: 72,
          width: format === 'rgb' ? 90 : 72, outline: 'none'
        }}
      />
      {open && (
        <div ref={popover} style={{
          position: 'absolute', top: 34, left: 0, zIndex: 1000,
          background: 'var(--bg-secondary)', border: '1px solid var(--border-color)',
          borderRadius: 'var(--radius)', padding: 10, boxShadow: '0 8px 24px rgba(0,0,0,0.5)'
        }}>
          <HexColorPicker color={hexValue} onChange={handleChange} />
        </div>
      )}
    </div>
  );
}
