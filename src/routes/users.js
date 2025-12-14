const express = require('express');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const User = require('../models/User');
const Wallet = require('../models/Wallet');
const userController = require('../controllers/users');
const { generateToken, generateEmailToken, verifyToken, requireRole } = require('../middleware/auth');
const { sendVerificationEmail, sendWelcomeEmail } = require('../services/mailService');
const router = express.Router();

router.post('/', async (req, res) => {
  try {
    const { role, name, email, password, agent_id } = req.body;
    
    const existingUser = await User.findOne({ email });
    if (existingUser) {
      return res.status(400).json({ error: 'User already exists with this email' });
    }
    
    const password_hash = await bcrypt.hash(password, 10);
    
    // Check if it's a test email - auto-verify test accounts
    const isTestEmail = email && email.toLowerCase().endsWith('@test.com');
    
    const user = new User({
      role: role || 'user',
      name,
      email,
      password_hash,
      agent_id,
      email_verified: isTestEmail // Auto-verify test emails
    });
    
    await user.save();
    
    const wallet = new Wallet({
      user_id: user._id
    });
    await wallet.save();
    
    // Only send verification email for non-test accounts
    if (!isTestEmail) {
      const emailToken = generateEmailToken(email);
      await sendVerificationEmail(email, emailToken);
    } else {
      console.log(`Test account created with auto-verification: ${email}`);
    }
    
    const userResponse = user.toObject();
    delete userResponse.password_hash;
    
    res.status(201).json({ 
      success: true,
      message: isTestEmail ? 'Test account created and verified.' : 'User created. Please verify email.',
      data: {
        user: userResponse, 
        wallet
      }
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

router.post('/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    
    const user = await User.findOne({ email });
    if (!user) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }
    
    const isValidPassword = await bcrypt.compare(password, user.password_hash);
    if (!isValidPassword) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }

    // Check if it's a test email
    const isTestAccount = email && email.toLowerCase().endsWith('@test.com');
    
    // Skip email verification for test accounts and in development
    if (!user.email_verified && !isTestAccount && process.env.NODE_ENV === 'production') {
      return res.status(401).json({ error: 'Please verify your email first' });
    }    const token = generateToken(user._id);
    
    const userResponse = user.toObject();
    delete userResponse.password_hash;
    
    res.json({ 
      success: true,
      message: 'Login successful',
      data: {
        user: userResponse, 
        token 
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

router.get('/verify-email', async (req, res) => {
  try {
    const { token } = req.query;
    const decoded = jwt.verify(token, process.env.EMAIL_SECRET || 'email-verification-key');
    
    const user = await User.findOneAndUpdate(
      { email: decoded.email },
      { email_verified: true },
      { new: true }
    );
    
    if (user) {
      await sendWelcomeEmail(user.email, user.name);
    }
    
    res.json({ message: 'Email verified successfully' });
  } catch (error) {
    res.status(400).json({ error: 'Invalid or expired token' });
  }
});

// Forgot Password - Request OTP
router.post('/forgot-password', async (req, res) => {
  try {
    const { email } = req.body;
    
    const user = await User.findOne({ email });
    if (!user) {
      // Don't reveal if email exists
      return res.json({ 
        success: true, 
        message: 'If the email exists, an OTP has been sent' 
      });
    }
    
    // Generate 6-digit OTP
    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const otpExpires = new Date(Date.now() + 10 * 60 * 1000); // 10 minutes
    
    user.reset_otp = otp;
    user.reset_otp_expires = otpExpires;
    await user.save();
    
    // Send OTP email
    const { sendPasswordResetOTP } = require('../services/mailService');
    await sendPasswordResetOTP(email, otp);
    
    res.json({ 
      success: true, 
      message: 'OTP sent to your email' 
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Verify OTP
router.post('/verify-otp', async (req, res) => {
  try {
    const { email, otp } = req.body;
    
    const user = await User.findOne({ 
      email,
      reset_otp: otp,
      reset_otp_expires: { $gt: new Date() }
    });
    
    if (!user) {
      return res.status(400).json({ error: 'Invalid or expired OTP' });
    }
    
    // Generate a temporary reset token
    const resetToken = jwt.sign(
      { userId: user._id, purpose: 'password-reset' },
      process.env.JWT_SECRET || 'your-secret-key',
      { expiresIn: '15m' }
    );
    
    res.json({ 
      success: true, 
      message: 'OTP verified',
      resetToken 
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Reset Password
router.post('/reset-password', async (req, res) => {
  try {
    const { resetToken, newPassword } = req.body;
    
    if (!resetToken || !newPassword) {
      return res.status(400).json({ error: 'Reset token and new password are required' });
    }
    
    if (newPassword.length < 6) {
      return res.status(400).json({ error: 'Password must be at least 6 characters' });
    }
    
    const decoded = jwt.verify(resetToken, process.env.JWT_SECRET || 'your-secret-key');
    
    if (decoded.purpose !== 'password-reset') {
      return res.status(400).json({ error: 'Invalid reset token' });
    }
    
    const user = await User.findById(decoded.userId);
    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }
    
    user.password_hash = await bcrypt.hash(newPassword, 10);
    user.reset_otp = undefined;
    user.reset_otp_expires = undefined;
    await user.save();
    
    res.json({ 
      success: true, 
      message: 'Password reset successfully' 
    });
  } catch (error) {
    if (error.name === 'TokenExpiredError') {
      return res.status(400).json({ error: 'Reset token has expired' });
    }
    res.status(500).json({ error: error.message });
  }
});

router.get('/', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const users = await User.find()
      .select('-password_hash')
      .populate('agent_id', 'name email');
    res.json({
      success: true,
      data: users
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Profile management routes
router.get('/profile', verifyToken, userController.getProfile);
router.put('/profile', verifyToken, userController.updateProfile);

// Agent referral routes
router.get('/referral/:code', async (req, res) => {
  try {
    const { code } = req.params;
    const agent = await User.findOne({ referral_code: code, role: 'agent', is_active: true })
      .select('name company_name referral_code');
    
    if (!agent) {
      return res.status(404).json({ error: 'Invalid referral code' });
    }
    
    res.json({
      success: true,
      data: {
        agent_id: agent._id,
        agent_name: agent.name,
        company_name: agent.company_name
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Agent: Get referred users
router.get('/my-referrals', verifyToken, async (req, res) => {
  try {
    if (req.user.role !== 'agent') {
      return res.status(403).json({ error: 'Access denied. Agents only.' });
    }
    
    const { page = 1, limit = 10 } = req.query;
    const skip = (parseInt(page) - 1) * parseInt(limit);
    
    const [referredUsers, total] = await Promise.all([
      User.find({ agent_id: req.user._id })
        .select('name email createdAt')
        .sort({ createdAt: -1 })
        .skip(skip)
        .limit(parseInt(limit)),
      User.countDocuments({ agent_id: req.user._id })
    ]);
    
    res.json({
      success: true,
      data: {
        users: referredUsers,
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

// Agent: Get dashboard stats
router.get('/agent-stats', verifyToken, async (req, res) => {
  try {
    if (req.user.role !== 'agent') {
      return res.status(403).json({ error: 'Access denied. Agents only.' });
    }
    
    const Commission = require('../models/Commission');
    const Withdrawal = require('../models/Withdrawal');
    const Report = require('../models/Report');
    
    // Get referred users count
    const referredUsersCount = await User.countDocuments({ agent_id: req.user._id });
    
    // Get commission stats
    const commissions = await Commission.find({ agent_id: req.user._id });
    const totalEarned = commissions.reduce((sum, c) => sum + c.commission_amount, 0);
    const accruedAmount = commissions.filter(c => c.status === 'accrued')
      .reduce((sum, c) => sum + c.commission_amount, 0);
    const paidAmount = commissions.filter(c => c.status === 'paid')
      .reduce((sum, c) => sum + c.commission_amount, 0);
    
    // Get pending withdrawals
    const pendingWithdrawals = await Withdrawal.find({
      agent_id: req.user._id,
      status: { $in: ['pending', 'approved', 'processing'] }
    });
    const pendingWithdrawalAmount = pendingWithdrawals.reduce((sum, w) => sum + w.amount, 0);
    
    // Get reports generated by referred users
    const referredUserIds = await User.find({ agent_id: req.user._id }).select('_id');
    const reportsFromReferrals = await Report.countDocuments({
      user_id: { $in: referredUserIds.map(u => u._id) }
    });
    
    res.json({
      success: true,
      data: {
        referral_code: req.user.referral_code,
        commission_rate: req.user.commission_rate,
        referred_users: referredUsersCount,
        reports_from_referrals: reportsFromReferrals,
        total_earned: totalEarned,
        available_balance: accruedAmount - pendingWithdrawalAmount,
        pending_withdrawal: pendingWithdrawalAmount,
        total_withdrawn: paidAmount
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Admin: Update agent commission rate
router.put('/:userId/commission-rate', verifyToken, async (req, res) => {
  try {
    if (req.user.role !== 'super_admin' && req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Access denied' });
    }
    
    const { userId } = req.params;
    const { commission_rate } = req.body;
    
    if (commission_rate < 0 || commission_rate > 100) {
      return res.status(400).json({ error: 'Commission rate must be between 0 and 100' });
    }
    
    const user = await User.findById(userId);
    if (!user || user.role !== 'agent') {
      return res.status(404).json({ error: 'Agent not found' });
    }
    
    user.commission_rate = commission_rate;
    await user.save();
    
    const userResponse = user.toObject();
    delete userResponse.password_hash;
    
    res.json({
      success: true,
      message: 'Commission rate updated',
      data: userResponse
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Admin: Update user (role, status, etc.)
router.patch('/:userId', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { userId } = req.params;
    const { role, status, is_active, name, email, mobile, bank_details, upi_details } = req.body;
    
    const updateData = {};
    if (role !== undefined) updateData.role = role;
    if (status !== undefined) updateData.status = status;
    if (is_active !== undefined) updateData.is_active = is_active;
    if (name !== undefined) updateData.name = name;
    if (email !== undefined) updateData.email = email;
    if (mobile !== undefined) updateData.mobile = mobile;
    if (bank_details !== undefined) updateData.bank_details = bank_details;
    if (upi_details !== undefined) updateData.upi_details = upi_details;
    
    const user = await User.findByIdAndUpdate(
      userId,
      updateData,
      { new: true }
    ).select('-password_hash');
    
    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }
    
    res.json({
      success: true,
      message: 'User updated successfully',
      data: user
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Admin: Delete user
router.delete('/:userId', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { userId } = req.params;
    
    const user = await User.findByIdAndDelete(userId);
    
    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }
    
    res.json({
      success: true,
      message: 'User deleted successfully'
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Admin: Toggle user active status
router.put('/:userId/status', verifyToken, async (req, res) => {
  try {
    if (req.user.role !== 'super_admin' && req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Access denied' });
    }
    
    const { userId } = req.params;
    const { is_active } = req.body;
    
    const user = await User.findByIdAndUpdate(
      userId,
      { is_active },
      { new: true }
    ).select('-password_hash');
    
    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }
    
    res.json({
      success: true,
      message: `User ${is_active ? 'enabled' : 'disabled'} successfully`,
      data: user
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Super admin routes
router.post('/create-super-admin', verifyToken, userController.createSuperAdmin);
router.get('/all', verifyToken, userController.getAllUsers);
router.put('/:userId/role', verifyToken, userController.updateUserRole);
router.put('/:userId/credits', verifyToken, userController.updateUserCredits);
router.delete('/:userId', verifyToken, userController.deleteUser);

module.exports = router;