const fs = require('fs').promises;
const path = require('path');
const config = require('../config/environment');

/**
 * Advanced Logging System
 * Production-ready logging with structured logs, rotation, and multiple log levels
 */

class Logger {
  constructor() {
    this.logsDir = path.join(__dirname, '../../logs');
    this.maxFileSize = 10 * 1024 * 1024; // 10MB
    this.maxFiles = 5; // Keep 5 rotated files
    this.initialized = false;
  }

  async initialize() {
    if (this.initialized) return;

    try {
      // Ensure logs directory exists
      await fs.mkdir(this.logsDir, { recursive: true });

      // Create log files if they don't exist
      const logFiles = [
        'app.log',
        'error.log',
        'access.log',
        'api.log',
        'debug.log'
      ];

      for (const file of logFiles) {
        const filePath = path.join(this.logsDir, file);
        try {
          await fs.access(filePath);
        } catch {
          await fs.writeFile(filePath, '');
        }
      }

      this.initialized = true;
      console.log('📝 Logger initialized successfully');
    } catch (error) {
      console.error('❌ Failed to initialize logger:', error);
    }
  }

  async _writeLog(level, message, data = {}, file = 'app.log') {
    if (!this.initialized) await this.initialize();

    const timestamp = new Date().toISOString();
    const logEntry = {
      timestamp,
      level: level.toUpperCase(),
      message,
      ...data
    };

    const logLine = JSON.stringify(logEntry) + '\n';
    const filePath = path.join(this.logsDir, file);

    try {
      // Check file size and rotate if needed
      await this._rotateFileIfNeeded(filePath);

      // Write to file
      await fs.appendFile(filePath, logLine);

      // Also write to console in development
      if (config.NODE_ENV === 'development') {
        const consoleMessage = `[${timestamp}] ${level.toUpperCase()}: ${message}`;
        if (level === 'error') {
          console.error(consoleMessage, data);
        } else if (level === 'warn') {
          console.warn(consoleMessage, data);
        } else {
          console.log(consoleMessage, data);
        }
      }
    } catch (error) {
      console.error('❌ Failed to write log:', error);
    }
  }

  async _rotateFileIfNeeded(filePath) {
    try {
      const stats = await fs.stat(filePath);
      if (stats.size > this.maxFileSize) {
        // Rotate files
        for (let i = this.maxFiles - 1; i >= 1; i--) {
          const oldFile = `${filePath}.${i}`;
          const newFile = `${filePath}.${i + 1}`;
          try {
            await fs.rename(oldFile, newFile);
          } catch {
            // File might not exist, continue
          }
        }

        // Move current file to .1
        await fs.rename(filePath, `${filePath}.1`);

        // Create new empty file
        await fs.writeFile(filePath, '');
      }
    } catch {
      // File might not exist yet, continue
    }
  }

  // Application logs
  info(message, data = {}) {
    return this._writeLog('info', message, data, 'app.log');
  }

  warn(message, data = {}) {
    return this._writeLog('warn', message, data, 'app.log');
  }

  error(message, data = {}, stack = null) {
    const logData = { ...data };
    if (stack) logData.stack = stack;
    return this._writeLog('error', message, logData, 'error.log');
  }

  debug(message, data = {}) {
    return this._writeLog('debug', message, data, 'debug.log');
  }

  // API specific logs
  apiRequest(req, res, responseTime) {
    const logData = {
      method: req.method,
      url: req.url,
      userAgent: req.get('User-Agent'),
      ip: req.ip,
      userId: req.user ? req.user._id : null,
      statusCode: res.statusCode,
      responseTime: `${responseTime}ms`,
      timestamp: new Date().toISOString()
    };

    return this._writeLog('info', 'API Request', logData, 'api.log');
  }

  apiError(req, error, statusCode = 500) {
    const logData = {
      method: req.method,
      url: req.url,
      userId: req.user ? req.user._id : null,
      error: error.message,
      stack: error.stack,
      statusCode
    };

    return this._writeLog('error', 'API Error', logData, 'api.log');
  }

  // Access logs
  access(message, data = {}) {
    return this._writeLog('info', message, data, 'access.log');
  }

  // Business logic logs
  business(operation, details = {}) {
    return this._writeLog('info', `Business: ${operation}`, details, 'app.log');
  }

  // Security logs
  security(event, details = {}) {
    return this._writeLog('warn', `Security: ${event}`, details, 'app.log');
  }

  // Performance logs
  performance(operation, duration, details = {}) {
    const logData = {
      operation,
      duration: `${duration}ms`,
      ...details
    };
    return this._writeLog('info', 'Performance', logData, 'app.log');
  }

  // Database logs
  database(operation, collection, details = {}) {
    const logData = {
      operation,
      collection,
      ...details
    };
    return this._writeLog('info', 'Database', logData, 'app.log');
  }

  // Credit/Wallet logs
  wallet(userId, operation, amount, balance, details = {}) {
    const logData = {
      userId,
      operation,
      amount,
      balance,
      ...details
    };
    return this._writeLog('info', 'Wallet', logData, 'app.log');
  }

  // Report generation logs
  report(userId, templateId, operation, details = {}) {
    const logData = {
      userId,
      templateId,
      operation,
      ...details
    };
    return this._writeLog('info', 'Report', logData, 'app.log');
  }

  // Cleanup logs
  cleanup(operation, count, details = {}) {
    const logData = {
      operation,
      count,
      ...details
    };
    return this._writeLog('info', 'Cleanup', logData, 'app.log');
  }

  // Get log files info
  async getLogFiles() {
    if (!this.initialized) await this.initialize();

    try {
      const files = await fs.readdir(this.logsDir);
      const logFiles = [];

      for (const file of files) {
        if (file.endsWith('.log')) {
          const filePath = path.join(this.logsDir, file);
          const stats = await fs.stat(filePath);
          logFiles.push({
            name: file,
            size: stats.size,
            modified: stats.mtime,
            path: filePath
          });
        }
      }

      return logFiles;
    } catch (error) {
      console.error('❌ Failed to get log files:', error);
      return [];
    }
  }

  // Read log file content
  async readLogFile(filename, lines = 100) {
    if (!this.initialized) await this.initialize();

    try {
      const filePath = path.join(this.logsDir, filename);
      const content = await fs.readFile(filePath, 'utf8');
      const logLines = content.trim().split('\n').filter(line => line.trim());

      // Return last N lines
      return logLines.slice(-lines);
    } catch (error) {
      console.error(`❌ Failed to read log file ${filename}:`, error);
      return [];
    }
  }
}

// Create singleton instance
const logger = new Logger();

// Initialize on module load
logger.initialize().catch(console.error);

module.exports = logger;