import { useEffect } from 'react';
import useThemeStore from '../../store/themeStore';

export default function StatusBar() {
  const bridgeStatus = useThemeStore((s) => s.bridgeStatus);
  const lastAction = useThemeStore((s) => s.lastAction);
  const setBridgeStatus = useThemeStore((s) => s.setBridgeStatus);

  useEffect(() => {
    if (!window.themeAPI?.onBridgeStatus) return;
    const unsubscribe = window.themeAPI.onBridgeStatus((status) => {
      setBridgeStatus(status);
    });
    return unsubscribe;
  }, [setBridgeStatus]);

  const statusLabel = {
    connecting: 'Connecting to PowerShell...',
    connected: 'PowerShell bridge connected',
    disconnected: 'PowerShell bridge disconnected',
    error: 'PowerShell bridge error'
  };

  return (
    <div className="statusbar">
      <div>
        <span className={`statusbar__dot statusbar__dot--${bridgeStatus}`} />
        {statusLabel[bridgeStatus] || bridgeStatus}
      </div>
      <div>
        {lastAction && (
          <span style={{ color: lastAction.success ? '#999' : 'var(--error)' }}>
            {lastAction.time} — {lastAction.text}
          </span>
        )}
      </div>
    </div>
  );
}
