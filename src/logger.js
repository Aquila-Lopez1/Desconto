// src/logger.js

/**
 * Logger module for centralized logging system.
 * Provides methods for logging errors and audit trails.
 */

class Logger {
    constructor() {
        this.logs = [];
    }

    logError(error) {
        const timestamp = new Date().toISOString();
        this.logs.push({
            timestamp,
            type: 'error',
            message: error.message,
        });
        console.error(`[ERROR] ${timestamp}: ${error.message}`);
    }

    logInfo(message) {
        const timestamp = new Date().toISOString();
        this.logs.push({
            timestamp,
            type: 'info',
            message,
        });
        console.log(`[INFO] ${timestamp}: ${message}`);
    }

    getLogs() {
        return this.logs;
    }
}

// Example usage
const logger = new Logger();
logger.logInfo('Logger initialized.');

export default logger;