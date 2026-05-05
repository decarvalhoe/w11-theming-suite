import { useEffect, useRef } from 'react';
import useThemeStore from '../store/themeStore';

/**
 * useLivePreview — Watches a specific theme section and applies changes
 * to the live system via IPC with debouncing.
 *
 * Uses a trailing debounce: each new change resets the timer so that
 * only the *last* value in a burst is sent to PowerShell.  While a
 * command is in-flight we hold the next value and send it once the
 * current call resolves — this prevents queue pile-up.
 *
 * @param {string} section  – The theme section key ('mode', 'accentColor', etc.)
 * @param {number} delay    – Debounce delay in ms (default 300)
 */
export function useLivePreview(section, delay = 300) {
  const sectionData = useThemeStore((s) => s.theme[section]);
  const setLastAction = useThemeStore((s) => s.setLastAction);
  const timerRef = useRef(null);
  const prevRef = useRef(null);
  const initialised = useRef(false);
  const inflightRef = useRef(false);   // true while a PS command is running
  const pendingRef = useRef(null);     // queued value to send after inflight resolves

  useEffect(() => {
    // Skip the initial render (don't apply on mount)
    if (!initialised.current) {
      prevRef.current = JSON.stringify(sectionData);
      initialised.current = true;
      return;
    }

    const current = JSON.stringify(sectionData);
    if (current === prevRef.current) return;
    prevRef.current = current;

    // If a PS command is already running, just store latest value — it will
    // be sent when the inflight command finishes.
    if (inflightRef.current) {
      pendingRef.current = sectionData;
      return;
    }

    if (timerRef.current) clearTimeout(timerRef.current);

    timerRef.current = setTimeout(() => {
      sendToPS(sectionData);
    }, delay);

    return () => {
      if (timerRef.current) clearTimeout(timerRef.current);
    };
  }, [sectionData, section, delay, setLastAction]);

  async function sendToPS(data) {
    if (!window.themeAPI) return;
    inflightRef.current = true;
    try {
      const sectionMap = {
        mode: 'DarkMode',
        accentColor: 'AccentColor',
        dwm: 'DWM',
        taskbar: 'Taskbar'
      };

      const psSection = sectionMap[section];

      if (psSection) {
        const config = { [section]: data };
        await window.themeAPI.applySection(psSection, config);
        setLastAction(`Applied ${psSection}`);
      } else if (section === 'win32Colors') {
        await window.themeAPI.applyWin32Colors(data);
        setLastAction('Applied Win32 Colors');
      }
    } catch (err) {
      setLastAction(`Error applying ${section}: ${err.message}`, false);
    } finally {
      inflightRef.current = false;

      // If a newer value arrived while we were busy, send it now
      if (pendingRef.current) {
        const next = pendingRef.current;
        pendingRef.current = null;
        prevRef.current = JSON.stringify(next);
        sendToPS(next);
      }
    }
  }
}

/**
 * useThemeInit — Load the current system theme state on first mount.
 */
export function useThemeInit() {
  const loadTheme = useThemeStore((s) => s.loadTheme);
  const setBridgeStatus = useThemeStore((s) => s.setBridgeStatus);

  useEffect(() => {
    const init = async () => {
      if (!window.themeAPI) {
        setBridgeStatus('disconnected');
        return;
      }
      try {
        const result = await window.themeAPI.readTheme();
        if (result.success && result.data) {
          // FIX C2: loadTheme expects a plain object, not a function.
          // defaultTheme is already merged inside loadTheme.
          loadTheme(result.data);
        }
        setBridgeStatus('connected');
      } catch (err) {
        console.error('Failed to read theme:', err);
        setBridgeStatus('error');
      }
    };
    init();
  }, [loadTheme, setBridgeStatus]);
}
