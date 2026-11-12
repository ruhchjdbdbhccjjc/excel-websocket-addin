/**
 * SettingsStorage - Wrapper around Office.js settings API
 * Provides type-safe storage for WebSocket connection settings
 */

export class SettingsStorage {
  private saveTimeout: NodeJS.Timeout | null = null;
  private saveDelay: number = 500; // Debounce save by 500ms

  private getOfficeSettings(): Office.Settings | null {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
        return Office.context.document.settings;
      }
    } catch {
      // ignore
    }
    return null;
  }

  /**
   * Get a setting value
   * @param key - The setting key
   * @param defaultValue - Default value if key doesn't exist
   * @returns Promise with the setting value
   */
  async get<T>(key: string, defaultValue?: T): Promise<T> {
    const settings = this.getOfficeSettings();
    if (settings) {
      try {
        const value = settings.get(key);
        if (value !== undefined) return value as T;
      } catch {
        // ignore
      }
    }

    // Graceful fallback to localStorage
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        const item = window.localStorage.getItem(`excel_addin_${key}`);
        if (item !== null) {
          return JSON.parse(item);
        }
      }
    } catch {
      // ignore
    }

    return defaultValue as T;
  }

  /**
   * Set a setting value
   * @param key - The setting key
   * @param value - The value to set
   * @returns Promise when set is complete
   */
  async set<T>(key: string, value: T): Promise<void> {
    const settings = this.getOfficeSettings();
    if (settings) {
      try {
        settings.set(key, value);
        this.scheduleSave();
      } catch {
        // ignore
      }
    }

    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.setItem(`excel_addin_${key}`, JSON.stringify(value));
      }
    } catch {
      // ignore
    }
  }

  /**
   * Remove a setting
   * @param key - The setting key to remove
   * @returns Promise when removal is complete
   */
  async remove(key: string): Promise<void> {
    const settings = this.getOfficeSettings();
    if (settings) {
      try {
        settings.remove(key);
        this.scheduleSave();
      } catch {
        // ignore
      }
    }

    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.removeItem(`excel_addin_${key}`);
      }
    } catch {
      // ignore
    }
  }

  /**
   * Save settings to Office document
   * @returns Promise when save is complete
   */
  async save(): Promise<void> {
    const settings = this.getOfficeSettings();
    if (!settings) return;
    return new Promise((resolve) => {
      try {
        settings.saveAsync(() => resolve());
      } catch {
        resolve();
      }
    });
  }

  /**
   * Schedule a debounced save
   */
  private scheduleSave(): void {
    if (this.saveTimeout) {
      clearTimeout(this.saveTimeout);
    }
    this.saveTimeout = setTimeout(() => {
      this.save().catch((error) => {
        console.error('Failed to save settings:', error);
      });
    }, this.saveDelay);
  }

  /**
   * Get multiple settings at once
   * @param keys - Array of setting keys
   * @param defaults - Default values object
   * @returns Promise with settings object
   */
  async getMultiple<T extends Record<string, unknown>>(
    keys: string[],
    defaults?: T
  ): Promise<T> {
    const result: Partial<T> = {};
    for (const key of keys) {
      result[key as keyof T] = await this.get(key, defaults?.[key as keyof T]);
    }
    return result as T;
  }

  /**
   * Set multiple settings at once
   * @param values - Object with key-value pairs
   * @returns Promise when set is complete
   */
  async setMultiple<T extends Record<string, unknown>>(values: T): Promise<void> {
    for (const [key, value] of Object.entries(values)) {
      await this.set(key, value);
    }
  }
}
