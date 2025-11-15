const express = require('express');
const Commission = require('../models/Commission');
const { verifyToken, requireRole } = require('../middleware/auth');
const router = express.Router();

router.get('/', verifyToken, async (req, res) => {
  try {
    let query = {};
    
    if (req.user.role === 'agent') {
      query = { agent_id: req.user._id };
    } else if (req.user.role === 'user') {
      return res.status(403).json({ error: 'Access denied' });
    }
    
    const commissions = await Commission.find(query)
      .populate('agent_id', 'name email')
      .populate('user_id', 'name email')
      .populate('order_id', 'pack_type credits amount_paid')
      .sort({ createdAt: -1 });
    
    const totalCommission = commissions.reduce((sum, commission) => {
      return sum + commission.commission_amount;
    }, 0);
    
    const accruedCommission = commissions
      .filter(c => c.status === 'accrued')
      .reduce((sum, commission) => sum + commission.commission_amount, 0);
    
    const paidCommission = commissions
      .filter(c => c.status === 'paid')
      .reduce((sum, commission) => sum + commission.commission_amount, 0);
    
    res.json({
      commissions,
      summary: {
        total: totalCommission,
        accrued: accruedCommission,
        paid: paidCommission
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

router.patch('/:commissionId/status', verifyToken, requireRole('admin'), async (req, res) => {
  try {
    const { commissionId } = req.params;
    const { status } = req.body;
    
    const commission = await Commission.findByIdAndUpdate(
      commissionId,
      { status },
      { new: true }
    )
      .populate('agent_id', 'name email')
      .populate('user_id', 'name email')
      .populate('order_id', 'pack_type credits amount_paid');
    
    if (!commission) {
      return res.status(404).json({ error: 'Commission not found' });
    }
    
    res.json(commission);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

module.exports = router;