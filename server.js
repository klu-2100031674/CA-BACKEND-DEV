require('dotenv').config();
const app = require('./src/app');
const { connectDatabase } = require('./src/config/database');
const config = require('./src/config/environment');
const excelCalculationService = require('./src/services/excelCalculationService');

/**
 * Server Entry Point
 */

// Handle uncaught exceptions
process.on('uncaughtException', (err) => {
  console.error('❌ UNCAUGHT EXCEPTION! Shutting down...');
  console.error(err.name, err.message);
  process.exit(1);
});

// Connect to database
connectDatabase();

// Start server
const server = app.listen(config.PORT, () => {
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
      console.log(`🧹 Cleaned up ${count} old temporary files`);
    }
  });

// Schedule periodic cleanup (every 6 hours)
setInterval(() => {
  excelCalculationService.cleanupTempFiles(config.TEMP_FILE_MAX_AGE_HOURS);
}, 6 * 60 * 60 * 1000);

// Handle unhandled promise rejections
process.on('unhandledRejection', (err) => {
  console.error('❌ UNHANDLED REJECTION! Shutting down...');
  console.error(err.name, err.message);
  server.close(() => {
    process.exit(1);
  });
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('👋 SIGTERM received. Shutting down gracefully...');
  server.close(() => {
    console.log('💤 Process terminated');
  });
});

process.on('SIGINT', () => {
  console.log('\n👋 SIGINT received. Shutting down gracefully...');
  server.close(() => {
    console.log('💤 Process terminated');
  });
});
