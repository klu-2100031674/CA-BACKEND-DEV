require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const corsOptions = require('./config/cors');
const { errorHandler, notFound } = require('./middleware/errorHandler');
const apiLogger = require('./middleware/apiLogger');
const routes = require('./routes');
const config = require('./config/environment');
const logger = require('./utils/logger');

/**
 * Express Application Setup
 */

const app = express();

// CORS
app.use(cors(corsOptions));

// API Logging Middleware (must be before body parsers)
app.use(apiLogger);

// Body Parser
app.use(express.json({ limit: config.MAX_FILE_SIZE }));
app.use(express.urlencoded({ extended: true, limit: config.MAX_FILE_SIZE }));

// Request logging in development (now handled by apiLogger)
// if (config.NODE_ENV === 'development') {
//   app.use((req, res, next) => {
//     console.log(`${new Date().toISOString()} - ${req.method} ${req.path}`);
//     next();
//   });
// }

// Static files (for temp and upload files)
app.use('/temp', express.static(path.join(__dirname, '../temp')));
app.use('/uploads', express.static(path.join(__dirname, '../uploads')));

// API Routes
app.use('/api', routes);

// Root endpoint
app.get('/', (req, res) => {
  logger.access('Root endpoint accessed', { ip: req.ip });
  res.json({
    message: 'CA Report Generation API',
    version: '2.0.0',
    status: 'running',
    endpoints: {
      health: '/api/health',
      docs: '/api/docs' // Future: API documentation
    }
  });
});

// 404 Handler
app.use(notFound);

// Error Handler
app.use(errorHandler);

module.exports = app;
