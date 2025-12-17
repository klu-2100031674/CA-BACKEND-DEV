/**
 * CORS Configuration for Production
 */

const logger = require('../utils/logger');

const corsOptions = {
  origin: function (origin, callback) {
    logger.debug('CORS check for origin', {
      operation: 'corsOriginCheck',
      origin
    });
    // Allow requests with no origin (mobile apps, Postman, etc.)
    if (!origin) return callback(null, true);

    // Default allowed origins
    const defaultOrigins = [
      "http://localhost:5173",
      "http://localhost:5174", // Add localhost:5174
      'https://ca-front-end-dev.onrender.com',
      'https://ca-front-end-dev.onrender.com/'
    ];

    // Get allowed origins from env or use defaults
    const allowedOrigins = process.env.ALLOWED_ORIGINS
      ? process.env.ALLOWED_ORIGINS.split(',').map(url => url.trim())
      : defaultOrigins;

    logger.debug('CORS origin validation', {
      operation: 'corsOriginCheck',
      origin,
      allowedOrigins
    });

    // In development OR when accessing from localhost (for local testing)
    if (process.env.NODE_ENV === 'development' || (origin && origin.includes('localhost'))) {
      const allowedDevOrigins = [
        'http://localhost:3000',
        'http://localhost:5173',
        'http://localhost:5174', // Add localhost:5174
        'http://127.0.0.1:3000',
        'http://127.0.0.1:5173',
        'http://127.0.0.1:5174'  // Add 127.0.0.1:5174
      ];
      if (allowedDevOrigins.includes(origin) || (origin && origin.includes('localhost'))) {
        logger.info('CORS allowed for localhost origin', {
          operation: 'corsOriginCheck',
          origin,
          nodeEnv: process.env.NODE_ENV
        });
        return callback(null, true);
      }
    }

    if (allowedOrigins.includes(origin)) {
      logger.info('CORS allowed', {
        operation: 'corsOriginCheck',
        origin
      });
      callback(null, true);
    } else {
      logger.warn('CORS blocked origin', {
        operation: 'corsOriginCheck',
        origin,
        allowedOrigins
      });
      callback(new Error(`Not allowed by CORS: ${origin}`));
    }
  },
  credentials: true,
  optionsSuccessStatus: 200,
  methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS', 'HEAD'],
  allowedHeaders: [
    'Content-Type', 
    'Authorization', 
    'X-Requested-With',
    'Accept',
    'Origin',
    'Cache-Control',
    'X-File-Name'
  ],
  exposedHeaders: ['Content-Disposition', 'X-Total-Count'],
  preflightContinue: false
};

module.exports = corsOptions;
