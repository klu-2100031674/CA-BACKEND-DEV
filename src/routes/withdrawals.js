const express = require('express');
const Withdrawal = require('../models/Withdrawal');
const Commission = require('../models/Commission');
const User = require('../models/User');
const { verifyToken, requireRole } = require('../middleware/auth');
const razorpayService = require('../services/razorpayService');
const router = express.Router();

// Get all withdrawals (admin) or agent's own withdrawals
router.get('/', verifyToken, async (req, res) => {
  try {
    let query = {};
    
    if (req.user.role === 'agent') {
      query = { agent_id: req.user._id };
    } else if (req.user.role !== 'super_admin' && req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Access denied' });
    }
    
    const { status, page = 1, limit = 10 } = req.query;
    if (status) {
      query.status = status;
    }
    
    const skip = (parseInt(page) - 1) * parseInt(limit);
    
    const [withdrawals, total] = await Promise.all([
      Withdrawal.find(query)
        .populate('agent_id', 'name email referral_code')
        .populate('processed_by', 'name email')
        .sort({ createdAt: -1 })
        .skip(skip)
        .limit(parseInt(limit)),
      Withdrawal.countDocuments(query)
    ]);
    
    // Calculate summary
    const summary = await Withdrawal.aggregate([
      { $match: req.user.role === 'agent' ? { agent_id: req.user._id } : {} },
      {
        $group: {
          _id: '$status',
          total: { $sum: '$amount' },
          count: { $sum: 1 }
        }
      }
    ]);
    
    const summaryObj = {
      pending: { amount: 0, count: 0 },
      approved: { amount: 0, count: 0 },
      completed: { amount: 0, count: 0 },
      rejected: { amount: 0, count: 0 }
    };
    
    summary.forEach(item => {
      if (summaryObj[item._id]) {
        summaryObj[item._id] = { amount: item.total, count: item.count };
      }
    });
    
    res.json({
      success: true,
      data: {
        withdrawals,
        summary: summaryObj,
        pagination: {
          page: parseInt(page),
          limit: parseInt(limit),
          total,
          pages: Math.ceil(total / parseInt(limit))
        }
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Agent: Create withdrawal request
router.post('/', verifyToken, requireRole('agent'), async (req, res) => {
  try {
    const { amount, payment_method, payment_details } = req.body;
    
    // Validate agent has sufficient commission balance in wallet
    const Wallet = require('../models/Wallet');
    const wallet = await Wallet.findOne({ user_id: req.user._id });
    const availableBalance = wallet?.commission_balance || 0;
    
    // Get pending withdrawal amount
    const pendingWithdrawals = await Withdrawal.find({
      agent_id: req.user._id,
      status: { $in: ['pending', 'approved', 'processing'] }
    });
    
    const pendingAmount = pendingWithdrawals.reduce((sum, w) => sum + w.amount, 0);
    
    if (amount > (availableBalance - pendingAmount)) {
      return res.status(400).json({ 
        error: 'Insufficient balance',
        available: availableBalance - pendingAmount
      });
    }
    
    if (amount < 100) {
      return res.status(400).json({ error: 'Minimum withdrawal amount is ₹100' });
    }
    
    const withdrawal = new Withdrawal({
      agent_id: req.user._id,
      amount,
      payment_method,
      payment_details: payment_details || req.user.bank_details
    });
    
    await withdrawal.save();
    
    res.status(201).json({
      success: true,
      message: 'Withdrawal request submitted successfully',
      data: withdrawal
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Admin: Update withdrawal status
router.patch('/:withdrawalId/status', verifyToken, requireRole(['super_admin', 'admin']), async (req, res) => {
  try {
    const { withdrawalId } = req.params;
    const { status, admin_remarks, transaction_id } = req.body;
    
    const withdrawal = await Withdrawal.findById(withdrawalId);
    
    if (!withdrawal) {
      return res.status(404).json({ error: 'Withdrawal not found' });
    }
    
    // Validate status transition
    const validTransitions = {
      'pending': ['approved', 'rejected', 'completed'],
      'approved': ['processing', 'rejected', 'completed'],
      'processing': ['completed', 'rejected'],
      'completed': [],
      'rejected': ['pending']
    };
    
    if (!validTransitions[withdrawal.status]?.includes(status)) {
      return res.status(400).json({ 
        error: `Cannot transition from ${withdrawal.status} to ${status}` 
      });
    }

    withdrawal.status = status;
    withdrawal.admin_remarks = admin_remarks || withdrawal.admin_remarks;
    withdrawal.processed_by = req.user._id;
    withdrawal.processed_at = new Date();
    
    if (transaction_id) {
      withdrawal.transaction_id = transaction_id;
    }

    // Handle approval - create Razorpay payout based on payment method
    if (status === 'approved') {
      try {
        // Get agent details for email
        const agent = await User.findById(withdrawal.agent_id);

        // Validate payment details
        if (!withdrawal.payment_details) {
          return res.status(400).json({ error: 'Payment details are required for approval' });
        }

        let payout;
        const { account_holder_name } = withdrawal.payment_details;

        if (!account_holder_name) {
          return res.status(400).json({ error: 'Account holder name is required' });
        }

        // Create payout based on payment method
        if (withdrawal.payment_method === 'upi') {
          const { upi_id } = withdrawal.payment_details;

          if (!upi_id) {
            return res.status(400).json({ error: 'UPI ID is required for UPI transfer' });
          }

          // Validate UPI ID format
          const upiRegex = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+$/;
          if (!upiRegex.test(upi_id)) {
            return res.status(400).json({ error: 'Invalid UPI ID format' });
          }

          // Create Razorpay UPI payout
          const payoutData = {
            amount: withdrawal.amount * 100, // Convert to paisa
            upi_id,
            account_holder_name,
            agent_email: agent.email,
            narration: `Commission withdrawal for ${agent.name}`
          };

          payout = await razorpayService.createUpiPayout(payoutData);

        } else if (withdrawal.payment_method === 'bank_transfer') {
          const { account_number, ifsc_code } = withdrawal.payment_details;

          if (!account_number || !ifsc_code) {
            return res.status(400).json({ error: 'Account number and IFSC code are required for bank transfer' });
          }

          // Create Razorpay bank payout
          const payoutData = {
            amount: withdrawal.amount * 100, // Convert to paisa
            account_number,
            ifsc_code,
            account_holder_name,
            agent_email: agent.email,
            narration: `Commission withdrawal for ${agent.name}`
          };

          payout = await razorpayService.createBankPayout(payoutData);

        } else {
          return res.status(400).json({ error: 'Invalid payment method' });
        }

        // Update withdrawal with payout details
        withdrawal.razorpay_payout_id = payout.id;
        withdrawal.razorpay_contact_id = payout.contact_id;
        withdrawal.razorpay_fund_account_id = payout.fund_account_id;
        withdrawal.payout_status = payout.status;

        // Store payment method specific details
        if (withdrawal.payment_method === 'upi') {
          withdrawal.upi_id = payout.upi_id;
        } else if (withdrawal.payment_method === 'bank_transfer') {
          withdrawal.account_number = payout.account_number;
          withdrawal.ifsc_code = payout.ifsc_code;
        }

      } catch (payoutError) {
        console.error('Razorpay UPI payout error:', payoutError);
        withdrawal.payout_failure_reason = payoutError.message;
        withdrawal.admin_remarks = (withdrawal.admin_remarks || '') + ` | UPI Payout failed: ${payoutError.message}`;
      }
    }
    
    // If completed, mark related commissions as paid and deduct from wallet
    if (status === 'completed') {
      await Commission.updateMany(
        { agent_id: withdrawal.agent_id, status: 'accrued' },
        { status: 'paid' }
      );
      
      // Deduct from agent's wallet
      const Wallet = require('../models/Wallet');
      const wallet = await Wallet.findOne({ user_id: withdrawal.agent_id });
      if (wallet) {
        wallet.commission_balance = Math.max(0, wallet.commission_balance - withdrawal.amount);
        await wallet.save();
      }
    }

    await withdrawal.save();

    // Store invoice details for completed withdrawals
    if (status === 'completed') {
      withdrawal.invoice_number = `INV-${withdrawal._id.toString().slice(-8).toUpperCase()}`;
      await withdrawal.save();
    }

    const updatedWithdrawal = await Withdrawal.findById(withdrawalId)
      .populate('agent_id', 'name email')
      .populate('processed_by', 'name email');
    
    res.json({
      success: true,
      message: `Withdrawal ${status} successfully`,
      data: updatedWithdrawal
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Agent: Get available balance for withdrawal
router.get('/balance', verifyToken, requireRole('agent'), async (req, res) => {
  try {
    // Get wallet balance
    const Wallet = require('../models/Wallet');
    const wallet = await Wallet.findOne({ user_id: req.user._id });
    const commissionBalance = wallet?.commission_balance || 0;
    
    // Get pending withdrawal amount
    const pendingWithdrawals = await Withdrawal.find({
      agent_id: req.user._id,
      status: { $in: ['pending', 'approved', 'processing'] }
    });
    
    const pendingAmount = pendingWithdrawals.reduce((sum, w) => sum + w.amount, 0);
    
    // Get total withdrawn amount
    const completedWithdrawals = await Withdrawal.find({
      agent_id: req.user._id,
      status: 'completed'
    });
    
    const totalWithdrawn = completedWithdrawals.reduce((sum, w) => sum + w.amount, 0);
    
    res.json({
      success: true,
      data: {
        total_earned: commissionBalance + totalWithdrawn,
        available_balance: commissionBalance - pendingAmount,
        pending_withdrawal: pendingAmount,
        total_withdrawn: totalWithdrawn
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Webhook: Handle Razorpay payout status updates
router.post('/webhook/payout', async (req, res) => {
  try {
    const signature = req.headers['x-razorpay-signature'];
    const body = JSON.stringify(req.body);

    // Verify webhook signature
    if (!razorpayService.verifyWebhookSignature(body, signature)) {
      return res.status(400).json({ error: 'Invalid signature' });
    }

    const { event, payload } = req.body;

    if (event === 'payout.processed') {
      const payout = payload.payout.entity;
      
      // Find withdrawal by payout ID
      const withdrawal = await Withdrawal.findOne({ razorpay_payout_id: payout.id });
      
      if (withdrawal) {
        withdrawal.payout_status = 'processed';
        withdrawal.transaction_id = payout.utr || payout.transaction_id;
        await withdrawal.save();
        
        console.log(`Payout processed for withdrawal ${withdrawal._id}`);
      }
    } else if (event === 'payout.failed') {
      const payout = payload.payout.entity;
      
      // Find withdrawal by payout ID
      const withdrawal = await Withdrawal.findOne({ razorpay_payout_id: payout.id });
      
      if (withdrawal) {
        withdrawal.payout_status = 'failed';
        withdrawal.payout_failure_reason = payout.failure_reason || 'Payout failed';
        await withdrawal.save();
        
        console.log(`Payout failed for withdrawal ${withdrawal._id}: ${withdrawal.payout_failure_reason}`);
      }
    }

    res.json({ status: 'ok' });
  } catch (error) {
    console.error('Webhook processing error:', error);
    res.status(500).json({ error: 'Webhook processing failed' });
  }
});

module.exports = router;
