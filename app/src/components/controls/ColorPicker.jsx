import { useState, useRef, useEffect, useCallback } from 'react';
import { HexColorPicker } from 'react-colorful';

/**
 * ColorPicker — Supports #RRGGBB hex and "R G B" (Win32) formats.
 *
 * @param {string} value     – Current value (hex "#RRGGBB" or "R G B")
 * @param {function} onChange – Callback with the new value in the same format
 * @param {'hex'|'rgb'} format – Output format (default: 'hex')
 */
export default function ColorPicker({ value, onChange, format = 'hex' }) {
  const [open, setOpen] = useState(false);
  const popover = useRef(null);

  // Convert "R G B" → "#RRGGBB"
  const toHex = useCallback((val) => {
    if (!val) return '#000000';
    if (val.startsWith('#')) return val;
    const parts = val.split(' ').map(Number);
    if (parts.length === 3) {
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

  const handleChange = useCallback((hex) => {
    if (format === 'rgb') {
      onChange(toRgbString(hex));
    } else {
      onChange(hex);
    }
  }, [format, onChange, toRgbString]);

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
      <span style={{
        fontSize: 12, fontFamily: 'var(--font-mono)', color: 'var(--text-secondary)', minWidth: 60
      }}>
        {format === 'rgb' ? value : hexValue}
      </span>
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
