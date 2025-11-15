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

const razorpay = new Razorpay({
  key_id: process.env.RAZORPAY_KEY_ID,
  key_secret: process.env.RAZORPAY_KEY_SECRET
});

router.post('/create', verifyToken, async (req, res) => {
  try {
    const { pack_type, credits, amount_paid, currency = 'INR' } = req.body;
    
    const razorpayOrder = await razorpay.orders.create({
      amount: amount_paid * 100,
      currency,
      receipt: `order_${Date.now()}`,
      payment_capture: 1
    });
    
    const order = new Order({
      user_id: req.user._id,
      pack_type,
      credits,
      amount_paid,
      currency,
      payment_info: {
        rzp_order_id: razorpayOrder.id
      },
      status: 'pending'
    });
    
    await order.save();
    
    res.status(201).json({
      order,
      razorpay_order: razorpayOrder,
      key_id: process.env.RAZORPAY_KEY_ID
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

router.post('/verify', verifyToken, async (req, res) => {
  try {
    const { razorpay_order_id, razorpay_payment_id, razorpay_signature } = req.body;
    
    const body = razorpay_order_id + '|' + razorpay_payment_id;
    const expectedSignature = crypto
      .createHmac('sha256', process.env.RAZORPAY_KEY_SECRET)
      .update(body.toString())
      .digest('hex');
    
    if (expectedSignature !== razorpay_signature) {
      return res.status(400).json({ error: 'Invalid signature' });
    }
    
    const payment = await razorpay.payments.fetch(razorpay_payment_id);
    
    const order = await Order.findOneAndUpdate(
      { 'payment_info.rzp_order_id': razorpay_order_id, user_id: req.user._id },
      {
        status: 'paid',
        'payment_info.rzp_payment_id': razorpay_payment_id,
        'payment_info.method': payment.method,
        'payment_info.captured': payment.captured
      },
      { new: true }
    );
    
    if (!order) {
      return res.status(404).json({ error: 'Order not found' });
    }
    
    const wallet = await Wallet.findOne({ user_id: req.user._id });
    if (wallet) {
      if (order.pack_type === 'report') {
        wallet.report_credits += order.credits;
      } else if (order.pack_type === 'enquiry') {
        wallet.enquiry_credits += order.credits;
      }
      await wallet.save();
    }
    
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
      
      if (order) {
        const wallet = await Wallet.findOne({ user_id: order.user_id });
        if (wallet) {
          if (order.pack_type === 'report') {
            wallet.report_credits += order.credits;
          } else if (order.pack_type === 'enquiry') {
            wallet.enquiry_credits += order.credits;
          }
          await wallet.save();
        }
      }
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
    
    if (status === 'paid' && order.status !== 'paid') {
      const wallet = await Wallet.findOne({ user_id: order.user_id });
      if (wallet) {
        if (order.pack_type === 'report') {
          wallet.report_credits += order.credits;
        } else if (order.pack_type === 'enquiry') {
          wallet.enquiry_credits += order.credits;
        }
        await wallet.save();
      }
    }
    
    res.json(order);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

module.exports = router;