const logger = require('../utils/logger');

/**
 * API Logging Middleware
 * Logs all API requests with performance metrics and user context
 */

const apiLogger = (req, res, next) => {
  const startTime = Date.now();

  // Log request start
  logger.apiRequest(req, res, 0); // Will be updated with actual response time

  // Override res.end to log response
  const originalEnd = res.end;
  res.end = function(...args) {
    const responseTime = Date.now() - startTime;

    // Update the request log with response time
    logger.apiRequest(req, res, responseTime);

    // Log slow requests
    if (responseTime > 5000) {
      logger.performance('Slow API Request', responseTime, {
        method: req.method,
        url: req.url,
        userId: req.user ? req.user._id : null,
        statusCode: res.statusCode
      });
    }

    // Log errors
    if (res.statusCode >= 400) {
      logger.apiError(req, new Error(`HTTP ${res.statusCode}`), res.statusCode);
    }

    // Call original end method
    originalEnd.apply(this, args);
  };

  next();
};

module.exports = apiLogger;