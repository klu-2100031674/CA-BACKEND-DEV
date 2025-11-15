const express = require('express');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const User = require('../models/User');
const Wallet = require('../models/Wallet');
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
    
    const user = new User({
      role: role || 'user',
      name,
      email,
      password_hash,
      agent_id
    });
    
    await user.save();
    
    const wallet = new Wallet({
      user_id: user._id
    });
    await wallet.save();
    
    const emailToken = generateEmailToken(email);
    await sendVerificationEmail(email, emailToken);
    
    const userResponse = user.toObject();
    delete userResponse.password_hash;
    
    res.status(201).json({ 
      success: true,
      message: 'User created. Please verify email.',
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

    // Skip email verification in development
    if (!user.email_verified && process.env.NODE_ENV === 'production') {
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

router.get('/', verifyToken, requireRole('admin'), async (req, res) => {
  try {
    const users = await User.find()
      .select('-password_hash')
      .populate('agent_id', 'name email');
    res.json(users);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;