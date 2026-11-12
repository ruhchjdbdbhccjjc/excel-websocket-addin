import { render } from 'solid-js/web';
import { App } from './App';
import { BackgroundService } from '../websocket/service/BackgroundService';
import { Logger } from '../websocket/logging/Logger';

// 0. Install global diagnostic hooks at the earliest possible lifecycle moment
Logger.getInstance().installGlobalHooks();

let isMounted = false;

function mountApp(): void {
  if (isMounted) return;
  const root = document.getElementById('root');
  if (!root) return;

  isMounted = true;
  root.innerHTML = ''; // Crucial: remove .loading-state placeholder

  try {
    render(() => <App />, root);
  } catch (err) {
    console.error('Failed to render SolidJS App:', err);
    root.innerHTML = `<div style="padding:16px;color:#f87171;font-family:sans-serif;">
      <h3>UI 加载错误</h3>
      <pre style="white-space:pre-wrap;font-size:12px;">${String(err)}</pre>
    </div>`;
  }
}

// 1. Mount immediately if DOM is ready
if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mountApp);
  } else {
    mountApp();
  }
}

// 2. Office lifecycle initialization
if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    mountApp();

    const isDialogMode = typeof window !== 'undefined' && Boolean(window.location?.search?.includes('dialog=true'));
    // In child dialog mode, DO NOT initialize BackgroundService (avoids AccessDenied on setStartupBehavior & duplicated WebSocket)
    if (info.host === Office.HostType.Excel && !isDialogMode) {
      try {
        const bgService = BackgroundService.getInstance();
        bgService.initialize().catch((err) => {
          console.warn('BackgroundService async init warning:', err);
        });
      } catch (err) {
        console.warn('BackgroundService sync error:', err);
      }
    }
  });
}
