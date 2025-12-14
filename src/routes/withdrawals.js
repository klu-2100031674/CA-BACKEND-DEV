const express = require('express');
const Withdrawal = require('../models/Withdrawal');
const Commission = require('../models/Commission');
const User = require('../models/User');
const { verifyToken, requireRole } = require('../middleware/auth');
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
    
    // Validate agent has sufficient commission balance
    const commissions = await Commission.find({ 
      agent_id: req.user._id, 
      status: 'accrued' 
    });
    
    const availableBalance = commissions.reduce((sum, c) => sum + c.commission_amount, 0);
    
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
      'pending': ['approved', 'rejected'],
      'approved': ['processing', 'rejected'],
      'processing': ['completed', 'rejected']
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
    
    // If completed, mark related commissions as paid
    if (status === 'completed') {
      await Commission.updateMany(
        { agent_id: withdrawal.agent_id, status: 'accrued' },
        { status: 'paid' }
      );
    }
    
    await withdrawal.save();
    
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
    // Get total accrued commissions
    const commissions = await Commission.find({ 
      agent_id: req.user._id, 
      status: 'accrued' 
    });
    
    const totalAccrued = commissions.reduce((sum, c) => sum + c.commission_amount, 0);
    
    // Get pending withdrawal amount
    const pendingWithdrawals = await Withdrawal.find({
      agent_id: req.user._id,
      status: { $in: ['pending', 'approved', 'processing'] }
    });
    
    const pendingAmount = pendingWithdrawals.reduce((sum, w) => sum + w.amount, 0);
    
    // Get total paid commissions
    const paidCommissions = await Commission.find({
      agent_id: req.user._id,
      status: 'paid'
    });
    
    const totalPaid = paidCommissions.reduce((sum, c) => sum + c.commission_amount, 0);
    
    res.json({
      success: true,
      data: {
        total_earned: totalAccrued + totalPaid,
        available_balance: totalAccrued - pendingAmount,
        pending_withdrawal: pendingAmount,
        total_withdrawn: totalPaid
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
