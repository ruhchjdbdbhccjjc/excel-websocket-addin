import { createSignal } from 'solid-js';

export type ThemeMode = 'auto' | 'light' | 'dark';
export type ResolvedTheme = 'light' | 'dark';

const STORAGE_KEY = 'excel_addin_theme_mode';

function isDarkColor(hex: string): boolean {
  if (!hex || typeof hex !== 'string') return false;
  let clean = hex.replace('#', '').trim();
  if (clean.length === 3) {
    clean = clean.split('').map((c) => c + c).join('');
  }
  if (clean.length !== 6) return false;
  const num = parseInt(clean, 16);
  if (isNaN(num)) return false;
  const r = (num >> 16) & 255;
  const g = (num >> 8) & 255;
  const b = num & 255;
  // Standard ITU-R BT.601 perceived luminance formula (< 128 is dark)
  const luminance = (r * 299 + g * 587 + b * 114) / 1000;
  return luminance < 128;
}

function detectExcelOrSystemTheme(): ResolvedTheme {
  try {
    // 1. Office.js officeTheme API
    if (typeof Office !== 'undefined' && Office.context && (Office.context as any).officeTheme) {
      const theme = (Office.context as any).officeTheme;
      const bg = theme.bodyBackgroundColor;
      if (bg) {
        return isDarkColor(bg) ? 'dark' : 'light';
      }
    }

    // 2. Fallback to OS system dark mode
    if (typeof window !== 'undefined' && window.matchMedia) {
      if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
        return 'dark';
      }
      return 'light';
    }
  } catch {
    // ignore
  }

  return 'dark'; // Default dark mode for modern slick feel
}

function getSavedThemeMode(): ThemeMode {
  try {
    if (typeof window !== 'undefined' && window.localStorage) {
      const saved = window.localStorage.getItem(STORAGE_KEY) as ThemeMode;
      if (saved === 'auto' || saved === 'light' || saved === 'dark') {
        return saved;
      }
    }
  } catch {
    // ignore
  }
  return 'auto';
}

const [themeMode, setThemeModeSignal] = createSignal<ThemeMode>(getSavedThemeMode());
const [resolvedTheme, setResolvedThemeSignal] = createSignal<ResolvedTheme>(
  themeMode() === 'auto' ? detectExcelOrSystemTheme() : (themeMode() as ResolvedTheme)
);

function applyThemeToDOM(theme: ResolvedTheme) {
  if (typeof document !== 'undefined') {
    document.body.setAttribute('data-theme', theme);
  }
}

// Initial apply
applyThemeToDOM(resolvedTheme());

// Safe Office theme change listener registration deferred until Office is ready
export function initOfficeThemeListener(): void {
  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(() => {
      try {
        const officeContext = Office.context as any;
        if (officeContext && officeContext.officeTheme) {
          // Re-detect once Office.context is populated
          if (themeMode() === 'auto') {
            const next = detectExcelOrSystemTheme();
            setResolvedThemeSignal(next);
            applyThemeToDOM(next);
          }

          officeContext.officeTheme.addHandlerAsync?.(
            (Office as any).EventType?.OfficeThemeChanged,
            () => {
              if (themeMode() === 'auto') {
                const next = detectExcelOrSystemTheme();
                setResolvedThemeSignal(next);
                applyThemeToDOM(next);
              }
            }
          );
        }
      } catch {
        // ignore if not supported by host
      }
    });
  }
}

// OS system color scheme listener
if (typeof window !== 'undefined' && window.matchMedia) {
  const mq = window.matchMedia('(prefers-color-scheme: dark)');
  mq.addEventListener?.('change', () => {
    if (themeMode() === 'auto') {
      const next = detectExcelOrSystemTheme();
      setResolvedThemeSignal(next);
      applyThemeToDOM(next);
    }
  });
}

// Automatically register Office listener
initOfficeThemeListener();

export const useTheme = () => {
  const setThemeMode = (mode: ThemeMode) => {
    setThemeModeSignal(mode);
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.setItem(STORAGE_KEY, mode);
      }
    } catch {
      // ignore
    }

    const nextResolved = mode === 'auto' ? detectExcelOrSystemTheme() : mode;
    setResolvedThemeSignal(nextResolved);
    applyThemeToDOM(nextResolved);
  };

  return {
    themeMode,
    resolvedTheme,
    setThemeMode
  };
};
