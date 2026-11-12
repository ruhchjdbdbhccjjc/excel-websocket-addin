/**
 * Logger - Modern Office Add-in Diagnostic & Logging Hub
 * Features:
 * 1. Global Singleton with transparent constructor fallback (No split loggers)
 * 2. 1000-entry Ring Buffer with zero memory leak risk in Shared Runtime
 * 3. Scoped Tagging: SYSTEM, WEBSOCKET, WS-IN, WS-OUT, EXCEL-API, SPOTLIGHT, SLOT, CONSOLE, UNCAUGHT
 * 4. RichApi.Error & OfficeExtension.Error deep structural inspection
 * 5. Global Console & Uncaught Exception Interception
 */

export enum LogLevel {
  Debug = 'debug',
  Info = 'info',
  Warn = 'warn',
  Error = 'error'
}

export type LogScope =
  | 'SYSTEM'
  | 'WEBSOCKET'
  | 'WS-IN'
  | 'WS-OUT'
  | 'EXCEL-API'
  | 'SPOTLIGHT'
  | 'SLOT'
  | 'CONSOLE'
  | 'UNCAUGHT';

export interface LogEntry {
  id: number;
  timestamp: Date;
  level: LogLevel;
  scope: LogScope;
  message: string;
  data?: any;
  error?: {
    name?: string;
    message?: string;
    code?: string;
    debugInfo?: any;
    stack?: string;
  };
}

export class Logger {
  private static instance: Logger | null = null;
  private currentLevel: LogLevel = LogLevel.Debug; // Preserve full diagnostics in ring buffer
  private logs: LogEntry[] = [];
  private readonly maxLogs: number = 1000;
  private nextId: number = 1;
  private isHookInstalled = false;
  private isInternalLogging = false;

  public static getInstance(): Logger {
    if (!Logger.instance) {
      Logger.instance = new Logger();
    }
    return Logger.instance;
  }

  constructor() {
    if (Logger.instance) {
      return Logger.instance;
    }
    Logger.instance = this;
  }

  public setLevel(level: LogLevel): void {
    this.currentLevel = level;
  }

  public getLevel(): LogLevel {
    return this.currentLevel;
  }

  private shouldLog(level: LogLevel): boolean {
    const levels = [LogLevel.Debug, LogLevel.Info, LogLevel.Warn, LogLevel.Error];
    return levels.indexOf(level) >= levels.indexOf(this.currentLevel);
  }

  public log(level: LogLevel, scope: LogScope | string, message: string, data?: any): void {
    if (!this.shouldLog(level)) {
      return;
    }

    let errorDetail: LogEntry['error'] | undefined;
    if (data && typeof data === 'object') {
      if (data instanceof Error || data.name === 'RichApi.Error' || data.code !== undefined) {
        errorDetail = {
          name: data.name,
          message: data.message,
          code: data.code,
          debugInfo: data.debugInfo,
          stack: data.stack
        };
      }
    }

    const entry: LogEntry = {
      id: this.nextId++,
      timestamp: new Date(),
      level,
      scope: (scope || 'SYSTEM') as LogScope,
      message: String(message || ''),
      data: errorDetail ? undefined : this.sanitizeData(data),
      error: errorDetail
    };

    this.logs.push(entry);
    if (this.logs.length > this.maxLogs) {
      this.logs.shift();
    }

    // Mirror to original console with formatted tag when not inside internal hook
    if (!this.isInternalLogging && typeof window !== 'undefined' && window.console) {
      this.isInternalLogging = true;
      try {
        const timeStr = entry.timestamp.toLocaleTimeString();
        const tag = `[${timeStr}] [${level.toUpperCase()}] [${entry.scope}]`;
        const content = errorDetail
          ? `${entry.message} ${JSON.stringify(errorDetail)}`
          : entry.data !== undefined
          ? `${entry.message} ${JSON.stringify(entry.data)}`
          : entry.message;

        if (level === LogLevel.Error) {
          console.error(`${tag} ${content}`);
        } else if (level === LogLevel.Warn) {
          console.warn(`${tag} ${content}`);
        } else {
          console.log(`${tag} ${content}`);
        }
      } catch {
        // Ignore mirror logging errors
      } finally {
        this.isInternalLogging = false;
      }
    }
  }

  public debug(message: string, data?: any, scope: LogScope = 'SYSTEM'): void {
    this.log(LogLevel.Debug, scope, message, data);
  }

  public info(message: string, data?: any, scope: LogScope = 'SYSTEM'): void {
    this.log(LogLevel.Info, scope, message, data);
  }

  public warn(message: string, data?: any, scope: LogScope = 'SYSTEM'): void {
    this.log(LogLevel.Warn, scope, message, data);
  }

  public error(message: string, data?: any, scope: LogScope = 'SYSTEM'): void {
    this.log(LogLevel.Error, scope, message, data);
  }

  public getLogs(): LogEntry[] {
    return [...this.logs];
  }

  public clearLogs(): void {
    this.logs = [];
  }

  /**
   * Installs global monkey-patches for window.console, window.onerror, and unhandledrejection
   */
  public installGlobalHooks(): void {
    if (this.isHookInstalled || typeof window === 'undefined') return;
    this.isHookInstalled = true;

    const originalConsole = {
      log: window.console.log,
      info: window.console.info,
      warn: window.console.warn,
      error: window.console.error,
      debug: window.console.debug
    };

    const captureConsole = (level: LogLevel, originalFn: Function, args: any[]) => {
      try {
        originalFn.apply(window.console, args);
      } catch {
        // Ignore console apply errors
      }

      if (this.isInternalLogging) return;

      this.isInternalLogging = true;
      try {
        const first = args[0];
        const rest = args.length > 1 ? args.slice(1) : undefined;
        let msg = '';
        if (typeof first === 'string') {
          msg = first;
        } else {
          try {
            msg = JSON.stringify(first);
          } catch {
            msg = String(first);
          }
        }
        this.log(level, 'CONSOLE', msg, rest);
      } catch {
        // Safeguard
      } finally {
        this.isInternalLogging = false;
      }
    };

    window.console.log = (...args: any[]) => captureConsole(LogLevel.Info, originalConsole.log, args);
    window.console.info = (...args: any[]) => captureConsole(LogLevel.Info, originalConsole.info, args);
    window.console.warn = (...args: any[]) => captureConsole(LogLevel.Warn, originalConsole.warn, args);
    window.console.error = (...args: any[]) => captureConsole(LogLevel.Error, originalConsole.error, args);
    window.console.debug = (...args: any[]) => captureConsole(LogLevel.Debug, originalConsole.debug, args);

    window.addEventListener('error', (event: ErrorEvent) => {
      this.log(LogLevel.Error, 'UNCAUGHT', `Uncaught Exception: ${event.message}`, {
        filename: event.filename,
        lineno: event.lineno,
        colno: event.colno,
        error: event.error
      });
    });

    window.addEventListener('unhandledrejection', (event: PromiseRejectionEvent) => {
      this.log(LogLevel.Error, 'UNCAUGHT', `Unhandled Promise Rejection: ${event.reason?.message || String(event.reason)}`, {
        reason: event.reason
      });
    });

    this.info('Global diagnostic console & uncaught error hooks installed', null, 'SYSTEM');
  }

  private sanitizeData(data: any): any {
    if (data === undefined || data === null) return undefined;
    if (typeof data !== 'object') return data;
    try {
      const cache = new Set();
      return JSON.parse(
        JSON.stringify(data, (key, value) => {
          if (typeof value === 'object' && value !== null) {
            if (cache.has(value)) {
              return '[Circular]';
            }
            cache.add(value);
          }
          return value;
        })
      );
    } catch {
      return String(data);
    }
  }
}
