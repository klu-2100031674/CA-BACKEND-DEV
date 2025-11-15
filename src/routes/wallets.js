const express = require('express');
const Wallet = require('../models/Wallet');
const { verifyToken, requireRole } = require('../middleware/auth');
const router = express.Router();

// Get current user's wallet
router.get('/me', verifyToken, async (req, res) => {
  try {
    const wallet = await Wallet.findOne({ user_id: req.user._id }).populate('user_id', 'name email');
    
    if (!wallet) {
      return res.status(404).json({ 
        success: false,
        error: 'Wallet not found',
        message: 'No wallet found for this user'
      });
    }
    
    res.json({
      success: true,
      data: wallet,
      message: 'Wallet retrieved successfully'
    });
  } catch (error) {
    res.status(500).json({ 
      success: false,
      error: error.message,
      message: 'Failed to retrieve wallet'
    });
  }
});

router.get('/:userId', verifyToken, async (req, res) => {
  try {
    const { userId } = req.params;
    
    if (req.user.role !== 'admin' && req.user._id.toString() !== userId) {
      return res.status(403).json({ 
        success: false,
        error: 'Access denied',
        message: 'You can only access your own wallet'
      });
    }
    
    const wallet = await Wallet.findOne({ user_id: userId }).populate('user_id', 'name email');
    
    if (!wallet) {
      return res.status(404).json({ 
        success: false,
        error: 'Wallet not found',
        message: 'No wallet found for this user'
      });
    }
    
    res.json({
      success: true,
      data: wallet,
      message: 'Wallet retrieved successfully'
    });
  } catch (error) {
    res.status(500).json({ 
      success: false,
      error: error.message,
      message: 'Failed to retrieve wallet'
    });
  }
});

router.patch('/:userId', verifyToken, requireRole('admin'), async (req, res) => {
  try {
    const { userId } = req.params;
    const { report_credits, enquiry_credits, notes } = req.body;
    
    const updateData = {};
    if (report_credits !== undefined) updateData.report_credits = report_credits;
    if (enquiry_credits !== undefined) updateData.enquiry_credits = enquiry_credits;
    if (notes !== undefined) updateData.notes = notes;
    
    const wallet = await Wallet.findOneAndUpdate(
      { user_id: userId },
      updateData,
      { new: true }
    ).populate('user_id', 'name email');
    
    if (!wallet) {
      return res.status(404).json({ error: 'Wallet not found' });
    }
    
    res.json(wallet);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

module.exports = router;