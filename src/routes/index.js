const express = require('express');
const router = express.Router();
const logger = require('../utils/logger');

/**
 * Main Routes Index
 * Aggregates all route modules
 */

// Import route modules
const userRoutes = require('./users');
const walletRoutes = require('./wallets');
const orderRoutes = require('./orders');
const reportRoutes = require('./reports'); // Changed from reportRoutes to reports
const commissionRoutes = require('./commissions');
const schemeEligibilityRoutes = require('./schemeEligibility');
const excelFilesRoutes = require('./excelFiles');
// const adminRoutes = require('./admin');

// API health check
router.get('/health', (req, res) => {
  logger.access('Health check requested', {
    ip: req.ip,
    userAgent: req.get('User-Agent')
  });

  res.json({
    success: true,
    message: 'API is running',
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV,
    uptime: process.uptime(),
    memory: process.memoryUsage()
  });
});

// Mount routes
router.use('/users', userRoutes);
router.use('/wallets', walletRoutes);
router.use('/orders', orderRoutes);
router.use('/reports', reportRoutes);
router.use('/commissions', commissionRoutes);
router.use('/scheme-eligibility', schemeEligibilityRoutes);
router.use('/excel-files', excelFilesRoutes);
// router.use('/admin', adminRoutes);

module.exports = router;
