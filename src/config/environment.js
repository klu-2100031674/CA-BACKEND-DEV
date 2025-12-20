/**
 * Environment Configuration
 */

const logger = require('../utils/logger');

const config = {
  // Server
  PORT: process.env.PORT || 3000,
  NODE_ENV: process.env.NODE_ENV || 'development',

  // Database
  MONGODB_URI: process.env.NODE_ENV === 'production'
    ? process.env.MONGODB_URI_PROD : process.env.MONGODB_URI,

  // JWT
  JWT_SECRET: process.env.JWT_SECRET || 'your-super-secret-jwt-key-change-this-in-production',
  JWT_EXPIRE: process.env.JWT_EXPIRE || '7d',
  EMAIL_SECRET: process.env.EMAIL_SECRET || 'your-email-verification-secret-key',

  // Razorpay
  RAZORPAY_KEY_ID: process.env.RAZORPAY_KEY_ID,
  RAZORPAY_KEY_SECRET: process.env.RAZORPAY_KEY_SECRET,
  RAZORPAY_WEBHOOK_SECRET: process.env.RAZORPAY_WEBHOOK_SECRET,
  RAZORPAY_ACCOUNT_NUMBER: process.env.RAZORPAY_ACCOUNT_NUMBER,

  // Email
  EMAIL_USER: process.env.EMAIL_USER,
  EMAIL_PASSWORD: process.env.EMAIL_PASSWORD,
  COMPANY_NAME: process.env.COMPANY_NAME || 'CA Report Generation',
  APP_NAME: process.env.APP_NAME || 'CA-Dev',

  // Frontend URL
  FRONTEND_URL: process.env.FRONTEND_URL || 'http://localhost:5173',

  // File Upload
  MAX_FILE_SIZE: process.env.MAX_FILE_SIZE || '100mb',

  // CORS
  ALLOWED_ORIGINS: process.env.ALLOWED_ORIGINS || 'http://localhost:5173,http://localhost:3000',

  // App Settings
  TEMP_FILE_MAX_AGE_HOURS: parseInt(process.env.TEMP_FILE_MAX_AGE_HOURS || '24', 10),
};

// Validation
if (config.NODE_ENV === 'production') {
  const required = ['MONGODB_URI', 'JWT_SECRET', 'EMAIL_SECRET'];
  const missing = required.filter(key => !config[key]);

  if (missing.length > 0) {
    logger.error('Missing required environment variables', {
      operation: 'environmentValidation',
      missingVariables: missing
    });
    process.exit(1);
  }
}

process.env.MONGODB_URI = config.MONGODB_URI;

module.exports = config;
