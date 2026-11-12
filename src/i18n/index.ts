import { createSignal } from 'solid-js';
import type { SupportedLocale, TranslationSchema } from './types';
import { zhCN } from './locales/zh-CN';
import { enUS } from './locales/en-US';

export type { SupportedLocale, TranslationSchema };

const STORAGE_KEY = 'excel_addin_locale';
const dictionaries: Record<SupportedLocale, TranslationSchema> = {
  'zh-CN': zhCN,
  'en-US': enUS
};

function detectInitialLocale(): SupportedLocale {
  try {
    if (typeof window !== 'undefined' && window.localStorage) {
      const saved = window.localStorage.getItem(STORAGE_KEY) as SupportedLocale;
      if (saved && dictionaries[saved]) {
        return saved;
      }
    }

    // Auto-detect from Office host environment
    if (typeof Office !== 'undefined' && Office.context && Office.context.displayLanguage) {
      const lang = Office.context.displayLanguage.toLowerCase();
      if (lang.startsWith('zh')) {
        return 'zh-CN';
      }
    }

    // Auto-detect from browser navigator
    if (typeof navigator !== 'undefined' && navigator.language) {
      const navLang = navigator.language.toLowerCase();
      if (navLang.startsWith('zh')) {
        return 'zh-CN';
      }
    }
  } catch {
    // Ignore detection errors
  }

  // Default to English if not Chinese
  return 'en-US';
}

const [currentLocale, setCurrentLocale] = createSignal<SupportedLocale>(detectInitialLocale());

export const useI18n = () => {
  const t = (): TranslationSchema => dictionaries[currentLocale()] || dictionaries['en-US'];

  const setLocale = (newLocale: SupportedLocale) => {
    if (!dictionaries[newLocale]) return;
    setCurrentLocale(newLocale);
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.setItem(STORAGE_KEY, newLocale);
      }
    } catch {
      // Ignore storage errors
    }
  };

  return {
    locale: currentLocale,
    setLocale,
    t
  };
};
