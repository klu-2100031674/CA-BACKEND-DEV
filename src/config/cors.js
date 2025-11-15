/**
 * CORS Configuration for Production
 */

const corsOptions = {
  origin: function (origin, callback) {
    console.log('CORS check for origin:', origin);
    // Allow requests with no origin (mobile apps, Postman, etc.)
    if (!origin) return callback(null, true);

    // Default allowed origins
    const defaultOrigins = [
      'https://ca-front-end-dev.onrender.com',
      'https://ca-front-end-dev.onrender.com/'
    ];

    // Get allowed origins from env or use defaults
    const allowedOrigins = process.env.ALLOWED_ORIGINS
      ? process.env.ALLOWED_ORIGINS.split(',').map(url => url.trim())
      : defaultOrigins;

    console.log(`Origin: ${origin}`);
    console.log(`Allowed origins: ${allowedOrigins.join(', ')}`);

    // In development, allow all localhost origins
    if (process.env.NODE_ENV === 'development') {
      const isLocalhost = origin.includes('localhost') || origin.includes('127.0.0.1');
      if (isLocalhost) {
        console.log(`✓ CORS allowed (development - localhost): ${origin}`);
        return callback(null, true);
      }
    }

    if (allowedOrigins.includes(origin)) {
      console.log(`✓ CORS allowed: ${origin}`);
      callback(null, true);
    } else {
      console.error(`✗ CORS blocked origin: ${origin}`);
      console.error(`Allowed origins: ${allowedOrigins.join(', ')}`);
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
