const express = require('express');
const router = express.Router();

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
// const adminRoutes = require('./admin');

// API health check
router.get('/health', (req, res) => {
  res.json({
    success: true,
    message: 'API is running',
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV
  });
});

// Mount routes
router.use('/users', userRoutes);
router.use('/wallets', walletRoutes);
router.use('/orders', orderRoutes);
router.use('/reports', reportRoutes);
router.use('/commissions', commissionRoutes);
// router.use('/admin', adminRoutes);

module.exports = router;
