/**
 * @deprecated This file contains legacy credit-based payment system
 * New system uses pay-per-report model via:
 * - POST /api/reports/create-payment-order (create order)
 * - POST /api/reports/:reportId/verify-payment (verify payment)
 * 
 * These endpoints are kept for backward compatibility with existing orders
 */
const express = require('express');
const Razorpay = require('razorpay');
const crypto = require('crypto');
const Order = require('../models/Order');
const Wallet = require('../models/Wallet');
const Commission = require('../models/Commission');
const User = require('../models/User');
const { verifyToken, requireRole } = require('../middleware/auth');
const { sendOrderConfirmation } = require('../services/mailService');
const router = express.Router();

// Initialize Razorpay only if credentials are provided
let razorpay = null;
if (process.env.RAZORPAY_KEY_ID && process.env.RAZORPAY_KEY_SECRET) {
  razorpay = new Razorpay({
    key_id: process.env.RAZORPAY_KEY_ID,
    key_secret: process.env.RAZORPAY_KEY_SECRET
  });
  console.log('✅ Razorpay initialized successfully');
} else {
  console.log('⚠️ Razorpay credentials not found - payment gateway disabled');
}

/**
 * @deprecated Use POST /api/reports/create-payment-order instead
 * This endpoint is for legacy credit-based purchases only
 */
router.post('/create', verifyToken, async (req, res) => {
  try {
    const { pack_type, credits, amount_paid, currency = 'INR' } = req.body;
    
    let razorpayOrder = null;
    
    // Create Razorpay order if available
    if (razorpay) {
      razorpayOrder = await razorpay.orders.create({
        amount: amount_paid * 100, // Convert to paise
        currency,
        receipt: `order_${Date.now()}`,
        payment_capture: 1
      });
    }
    
    const order = new Order({
      user_id: req.user._id,
      pack_type,
      credits,
      amount_paid,
      currency,
      payment_info: {
        rzp_order_id: razorpayOrder?.id || null
      },
      status: 'pending'
    });
    
    await order.save();
    
    res.status(201).json({
      order,
      razorpay_order: razorpayOrder,
      key_id: process.env.RAZORPAY_KEY_ID || null
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * @deprecated Use POST /api/reports/:reportId/verify-payment instead
 * This endpoint is for legacy credit-based purchases only
 */
router.post('/verify', verifyToken, async (req, res) => {
  try {
    const { razorpay_order_id, razorpay_payment_id, razorpay_signature } = req.body;
    
    // Verify signature
    const body = razorpay_order_id + '|' + razorpay_payment_id;
    const expectedSignature = crypto
      .createHmac('sha256', process.env.RAZORPAY_KEY_SECRET)
      .update(body.toString())
      .digest('hex');
    
    if (expectedSignature !== razorpay_signature) {
      return res.status(400).json({ error: 'Invalid signature' });
    }
    
    // Fetch payment details if razorpay is available
    let paymentDetails = {};
    if (razorpay) {
      try {
        const payment = await razorpay.payments.fetch(razorpay_payment_id);
        paymentDetails = {
          method: payment.method,
          captured: payment.captured
        };
      } catch (err) {
        console.error('Failed to fetch payment details:', err);
      }
    }
    
    const order = await Order.findOneAndUpdate(
      { 'payment_info.rzp_order_id': razorpay_order_id, user_id: req.user._id },
      {
        status: 'paid',
        'payment_info.rzp_payment_id': razorpay_payment_id,
        'payment_info.method': paymentDetails.method,
        'payment_info.captured': paymentDetails.captured
      },
      { new: true }
    );
    
    if (!order) {
      return res.status(404).json({ error: 'Order not found' });
    }
    
    // Note: Wallet credit addition removed in pay-per-report model
    // Payment is now linked directly to reports
    
    if (req.user.agent_id) {
      const agent = await User.findById(req.user.agent_id);
      if (agent && agent.role === 'agent') {
        const commission = new Commission({
          agent_id: req.user.agent_id,
          user_id: req.user._id,
          order_id: order._id,
          rate_percent: 10,
          base_amount: order.amount_paid,
          commission_amount: (order.amount_paid * 10) / 100,
          status: 'accrued'
        });
        await commission.save();
      }
    }
    
    await sendOrderConfirmation(req.user.email, req.user.name, order);
    
    res.json({ message: 'Payment verified successfully', order });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

router.post('/webhook', async (req, res) => {
  try {
    const webhookSignature = req.headers['x-razorpay-signature'];
    const body = JSON.stringify(req.body);
    
    const expectedSignature = crypto
      .createHmac('sha256', process.env.RAZORPAY_WEBHOOK_SECRET)
      .update(body)
      .digest('hex');
    
    if (webhookSignature !== expectedSignature) {
      return res.status(400).json({ error: 'Invalid webhook signature' });
    }
    
    const event = req.body.event;
    const paymentEntity = req.body.payload.payment.entity;
    
    if (event === 'payment.captured') {
      const order = await Order.findOneAndUpdate(
        { 'payment_info.rzp_order_id': paymentEntity.order_id },
        {
          status: 'paid',
          'payment_info.rzp_payment_id': paymentEntity.id,
          'payment_info.method': paymentEntity.method,
          'payment_info.captured': true
        },
        { new: true }
      );
      
      // Note: Wallet credit addition removed in pay-per-report model
    } else if (event === 'payment.failed') {
      await Order.findOneAndUpdate(
        { 'payment_info.rzp_order_id': paymentEntity.order_id },
        { status: 'failed' }
      );
    }
    
    res.json({ status: 'ok' });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

router.get('/', verifyToken, async (req, res) => {
  try {
    const query = req.user.role === 'admin' ? {} : { user_id: req.user._id };
    
    const orders = await Order.find(query)
      .populate('user_id', 'name email')
      .sort({ createdAt: -1 });
    
    res.json(orders);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

router.patch('/:orderId/status', verifyToken, requireRole('admin'), async (req, res) => {
  try {
    const { orderId } = req.params;
    const { status } = req.body;
    
    const order = await Order.findByIdAndUpdate(
      orderId,
      { status },
      { new: true }
    ).populate('user_id', 'name email');
    
    if (!order) {
      return res.status(404).json({ error: 'Order not found' });
    }
    
    // Note: Wallet credit addition removed in pay-per-report model
    
    res.json(order);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

module.exports = router;