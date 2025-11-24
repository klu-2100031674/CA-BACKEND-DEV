require('dotenv').config();
const app = require('./src/app');
const { connectDatabase } = require('./src/config/database');
const config = require('./src/config/environment');
const excelCalculationService = require('./src/services/excelCalculationService');
const logger = require('./src/utils/logger');

/**
 * Server Entry Point
 */

// Handle uncaught exceptions
process.on('uncaughtException', (err) => {
  logger.error('Uncaught Exception', { error: err.message, stack: err.stack });
  console.error('❌ UNCAUGHT EXCEPTION! Shutting down...');
  console.error(err.name, err.message);
  process.exit(1);
});

// Connect to database
connectDatabase();

// Start server
const server = app.listen(config.PORT,'0.0.0.0', () => {
  logger.info('Server started successfully', {
    port: config.PORT,
    environment: config.NODE_ENV,
    nodeVersion: process.version
  });

  console.log('='.repeat(60));
  console.log('🚀 CA Report Generation API Server');
  console.log('='.repeat(60));
  console.log(`📡 Port: ${config.PORT}`);
  console.log(`🌍 Environment: ${config.NODE_ENV}`);
  console.log(`🔗 API URL: http://localhost:${config.PORT}`);
  console.log(`🏥 Health Check: http://localhost:${config.PORT}/api/health`);
  console.log('='.repeat(60));
});

// Cleanup temp files on startup
excelCalculationService.cleanupTempFiles(config.TEMP_FILE_MAX_AGE_HOURS)
  .then((count) => {
    if (count > 0) {
      logger.cleanup('Startup temp file cleanup', count, {
        maxAgeHours: config.TEMP_FILE_MAX_AGE_HOURS
      });
      console.log(`🧹 Cleaned up ${count} old temporary files`);
    }
  })
  .catch((error) => {
    logger.error('Failed to cleanup temp files on startup', { error: error.message });
  });

// Schedule periodic cleanup (every 6 hours)
setInterval(() => {
  excelCalculationService.cleanupTempFiles(config.TEMP_FILE_MAX_AGE_HOURS)
    .then((count) => {
      if (count > 0) {
        logger.cleanup('Periodic temp file cleanup', count, {
          maxAgeHours: config.TEMP_FILE_MAX_AGE_HOURS
        });
      }
    })
    .catch((error) => {
      logger.error('Failed to cleanup temp files periodically', { error: error.message });
    });
}, 6 * 60 * 60 * 1000);

// Handle unhandled promise rejections
process.on('unhandledRejection', (err) => {
  logger.error('Unhandled Promise Rejection', {
    error: err.message,
    stack: err.stack
  });
  console.error('❌ UNHANDLED REJECTION! Shutting down...');
  console.error(err.name, err.message);
  server.close(() => {
    logger.info('Server shut down gracefully due to unhandled rejection');
    process.exit(1);
  });
});

// Graceful shutdown
process.on('SIGTERM', () => {
  logger.info('SIGTERM received - initiating graceful shutdown');
  console.log('👋 SIGTERM received. Shutting down gracefully...');
  server.close(() => {
    logger.info('Server shut down gracefully');
    console.log('💤 Process terminated');
  });
});

process.on('SIGINT', () => {
  logger.info('SIGINT received - initiating graceful shutdown');
  console.log('\n👋 SIGINT received. Shutting down gracefully...');
  server.close(() => {
    logger.info('Server shut down gracefully');
    console.log('💤 Process terminated');
  });
});
