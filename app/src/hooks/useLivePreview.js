import { useEffect, useRef } from 'react';
import useThemeStore from '../store/themeStore';

/**
 * useLivePreview — Watches a specific theme section and applies changes
 * to the live system via IPC with debouncing.
 *
 * @param {string} section  – The theme section key ('mode', 'accentColor', etc.)
 * @param {number} delay    – Debounce delay in ms (default 150)
 */
export function useLivePreview(section, delay = 150) {
  const sectionData = useThemeStore((s) => s.theme[section]);
  const setLastAction = useThemeStore((s) => s.setLastAction);
  const timerRef = useRef(null);
  const prevRef = useRef(null);
  const initialised = useRef(false);

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

    if (timerRef.current) clearTimeout(timerRef.current);

    timerRef.current = setTimeout(async () => {
      if (!window.themeAPI) return;
      try {
        // Map section names to PS section parameter names
        const sectionMap = {
          mode: 'DarkMode',
          accentColor: 'AccentColor',
          dwm: 'DWM',
          taskbar: 'Taskbar'
        };

        const psSection = sectionMap[section];

        if (psSection) {
          // Standard registry sections
          const config = { [section]: sectionData };
          await window.themeAPI.applySection(psSection, config);
          setLastAction(`Applied ${psSection}`);
        } else if (section === 'win32Colors') {
          await window.themeAPI.applyWin32Colors(sectionData);
          setLastAction('Applied Win32 Colors');
        }
      } catch (err) {
        setLastAction(`Error applying ${section}: ${err.message}`, false);
      }
    }, delay);

    return () => {
      if (timerRef.current) clearTimeout(timerRef.current);
    };
  }, [sectionData, section, delay, setLastAction]);
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
