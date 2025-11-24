const mongoose = require('mongoose');
const logger = require('../utils/logger');

/**
 * Database Configuration
 */

const connectDatabase = async () => {
  try {
    // Set mongoose options
    mongoose.set('strictQuery', false);
    mongoose.set('bufferTimeoutMS', 30000); // 30 seconds timeout

    const options = {
      serverSelectionTimeoutMS: 30000,
      socketTimeoutMS: 45000,
      connectTimeoutMS: 30000,
      maxPoolSize: 10,
      minPoolSize: 5,
    };

    const conn = await mongoose.connect(process.env.MONGODB_URI, options);

    logger.info('MongoDB connected successfully', {
      operation: 'connectDatabase',
      host: conn.connection.host,
      database: conn.connection.name
    });

    // Handle connection events
    mongoose.connection.on('error', (err) => {
      logger.error('MongoDB connection error', {
        operation: 'databaseConnection',
        error: err.message,
        stack: err.stack
      });
    });

    mongoose.connection.on('disconnected', () => {
      logger.warn('MongoDB disconnected', {
        operation: 'databaseConnection'
      });
    });

    mongoose.connection.on('reconnected', () => {
      logger.info('MongoDB reconnected', {
        operation: 'databaseConnection'
      });
    });

  } catch (error) {
    logger.error('MongoDB connection failed', { operation: 'connectDatabase', error: error.message });
    process.exit(1);
  }
};

module.exports = { connectDatabase };
