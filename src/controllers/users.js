const User = require('../models/User');
const Wallet = require('../models/Wallet');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const logger = require('../utils/logger');
const r2Service = require('../services/cloudflareR2Service');

/**
 * User Controller
 * Handles user profile management and super admin functionality
 */

/**
 * @route   GET /api/users/profile
 * @desc    Get current user profile
 * @access  Private
 */
exports.getProfile = async (req, res, next) => {
  try {
    const user = await User.findById(req.user._id).select('-password_hash');

    if (!user) {
      return res.status(404).json({
        success: false,
        error: 'User not found'
      });
    }

    // Generate referral code for agents if it doesn't exist
    if (user.role === 'agent' && !user.referral_code) {
      user.referral_code = `AG${user.name.substring(0, 3).toUpperCase()}${Date.now().toString(36).toUpperCase()}`;
      await user.save();
      logger.info('Generated referral code for agent', {
        userId: req.user._id,
        referralCode: user.referral_code
      });
    }

    logger.info('Get profile response', {
      userId: req.user._id,
      role: user.role,
      hasReferralCode: !!user.referral_code,
      referralCode: user.referral_code,
      hasBankDetails: !!user.bank_details,
      bankDetailsKeys: user.bank_details ? Object.keys(user.bank_details) : null,
      bankDetails: user.bank_details
    });

    // Convert to object to modify
    const userObj = user.toObject();

    // Generate presigned URLs for images if they exist
    if (userObj.signature_url) {
      const key = r2Service.extractKeyFromUrl(userObj.signature_url);
      if (key) {
        userObj.signature_url = await r2Service.generatePresignedUrl(key);
      }
    }
    if (userObj.profile_logo) {
      const key = r2Service.extractKeyFromUrl(userObj.profile_logo);
      if (key) {
        userObj.profile_logo = await r2Service.generatePresignedUrl(key);
      }
    }
    if (userObj.company_logo) {
      const key = r2Service.extractKeyFromUrl(userObj.company_logo);
      if (key) {
        userObj.company_logo = await r2Service.generatePresignedUrl(key);
      }
    }

    res.json({
      success: true,
      data: userObj
    });

  } catch (error) {
    logger.error('Error in getProfile', {
      userId: req.user._id,
      error: error.message,
      stack: error.stack,
      operation: 'getProfile'
    });
    next(error);
  }
};

/**
 * @route   PUT /api/users/profile
 * @desc    Update user profile
 * @access  Private
 */
exports.updateProfile = async (req, res, next) => {
  try {
    const { name, company_name, phone, address, profile_logo, company_logo, signature_url, mobile, bank_details, upi_details } = req.body;

    logger.info('Update profile request', {
      userId: req.user._id,
      hasBankDetails: !!bank_details,
      bankDetailsKeys: bank_details ? Object.keys(bank_details) : null,
      bankDetails: bank_details
    });

    const updateData = {};
    if (name !== undefined) updateData.name = name;
    if (company_name !== undefined) updateData.company_name = company_name;
    if (phone !== undefined) updateData.phone = phone;
    if (address !== undefined) updateData.address = address;
    if (profile_logo !== undefined && profile_logo.startsWith('data:image')) {
      const buffer = Buffer.from(profile_logo.split(',')[1], 'base64');
      const fileName = `profile_${req.user._id}_${Date.now()}.png`;
      const contentType = profile_logo.split(';')[0].split(':')[1];
      updateData.profile_logo = await r2Service.uploadImage({
        fileBuffer: buffer,
        userEmail: req.user.email,
        fileName: fileName,
        contentType: contentType
      });
    } else if (profile_logo !== undefined) {
      updateData.profile_logo = profile_logo;
    }

    if (company_logo !== undefined && company_logo.startsWith('data:image')) {
      const buffer = Buffer.from(company_logo.split(',')[1], 'base64');
      const fileName = `company_${req.user._id}_${Date.now()}.png`;
      const contentType = company_logo.split(';')[0].split(':')[1];
      updateData.company_logo = await r2Service.uploadImage({
        fileBuffer: buffer,
        userEmail: req.user.email,
        fileName: fileName,
        contentType: contentType
      });
    } else if (company_logo !== undefined) {
      updateData.company_logo = company_logo;
    }

    if (signature_url !== undefined && signature_url.startsWith('data:image')) {
      const buffer = Buffer.from(signature_url.split(',')[1], 'base64');
      const fileName = `signature_${req.user._id}_${Date.now()}.png`;
      const contentType = signature_url.split(';')[0].split(':')[1];
      updateData.signature_url = await r2Service.uploadImage({
        fileBuffer: buffer,
        userEmail: req.user.email,
        fileName: fileName,
        contentType: contentType
      });
    } else if (signature_url !== undefined) {
      updateData.signature_url = signature_url;
    }

    if (mobile !== undefined) updateData.mobile = mobile;
    if (bank_details !== undefined) {
      // Merge upi_details into bank_details if it exists for backward compatibility
      updateData.bank_details = { ...bank_details };
      if (upi_details && upi_details.upi_id) {
        updateData.bank_details.upi_id = upi_details.upi_id;
      }
    }
    if (upi_details !== undefined && !bank_details) updateData.upi_details = upi_details;

    const user = await User.findByIdAndUpdate(
      req.user._id,
      updateData,
      { new: true, runValidators: true }
    ).select('-password_hash');

    if (!user) {
      return res.status(404).json({
        success: false,
        error: 'User not found'
      });
    }

    // Convert to object to modify
    const userObj = user.toObject();

    // Generate presigned URLs for images if they exist
    if (userObj.signature_url) {
      const key = r2Service.extractKeyFromUrl(userObj.signature_url);
      if (key) {
        userObj.signature_url = await r2Service.generatePresignedUrl(key);
      }
    }
    if (userObj.profile_logo) {
      const key = r2Service.extractKeyFromUrl(userObj.profile_logo);
      if (key) {
        userObj.profile_logo = await r2Service.generatePresignedUrl(key);
      }
    }
    if (userObj.company_logo) {
      const key = r2Service.extractKeyFromUrl(userObj.company_logo);
      if (key) {
        userObj.company_logo = await r2Service.generatePresignedUrl(key);
      }
    }

    res.json({
      success: true,
      message: 'Profile updated successfully',
      data: userObj
    });

  } catch (error) {
    logger.error('Error in updateProfile', {
      userId: req.user._id,
      error: error.message,
      stack: error.stack,
      operation: 'updateProfile'
    });
    next(error);
  }
};

/**
 * @route   POST /api/users/create-super-admin
 * @desc    Create a super admin user
 * @access  Private (Super Admin only)
 */
exports.createSuperAdmin = async (req, res, next) => {
  try {
    // Check if current user is super admin
    if (req.user.role !== 'super_admin') {
      return res.status(403).json({
        success: false,
        error: 'Only super admins can create other super admins'
      });
    }

    const { name, email, password, company_name, phone } = req.body;

    // Validate required fields
    if (!name || !email || !password) {
      return res.status(400).json({
        success: false,
        error: 'Name, email, and password are required'
      });
    }

    // Check if user already exists
    const existingUser = await User.findOne({ email });
    if (existingUser) {
      return res.status(400).json({
        success: false,
        error: 'User with this email already exists'
      });
    }

    // Hash password
    const saltRounds = 10;
    const password_hash = await bcrypt.hash(password, saltRounds);

    // Create super admin user
    const user = new User({
      role: 'super_admin',
      name,
      email,
      password_hash,
      company_name,
      phone,
      email_verified: true // Super admins are auto-verified
    });

    await user.save();

    // Create wallet for the super admin
    const wallet = new Wallet({
      user_id: user._id,
      report_credits: 1000, // Super admins get 1000 credits
      enquiry_credits: 1000
    });

    await wallet.save();

    logger.business('Super admin created successfully', {
      userId: user._id,
      email: user.email,
      name: user.name,
      role: user.role,
      creditsGranted: 1000,
      operation: 'createSuperAdmin'
    });

    res.status(201).json({
      success: true,
      message: 'Super admin created successfully',
      data: {
        _id: user._id,
        name: user.name,
        email: user.email,
        role: user.role,
        company_name: user.company_name,
        phone: user.phone,
        createdAt: user.createdAt
      }
    });

  } catch (error) {
    logger.error('Error in createSuperAdmin', {
      error: error.message,
      stack: error.stack,
      operation: 'createSuperAdmin'
    });
    next(error);
  }
};

/**
 * @route   GET /api/users
 * @desc    Get all users (Super Admin only)
 * @access  Private (Super Admin only)
 */
exports.getAllUsers = async (req, res, next) => {
  try {
    // Check if current user is super admin
    if (req.user.role !== 'super_admin') {
      return res.status(403).json({
        success: false,
        error: 'Access denied. Super admin privileges required.'
      });
    }

    const { page = 1, limit = 10, role, search } = req.query;

    let query = {};

    if (role) {
      query.role = role;
    }

    if (search) {
      query.$or = [
        { name: { $regex: search, $options: 'i' } },
        { email: { $regex: search, $options: 'i' } },
        { company_name: { $regex: search, $options: 'i' } }
      ];
    }

    const users = await User.find(query)
      .select('-password_hash')
      .sort({ createdAt: -1 })
      .limit(limit * 1)
      .skip((page - 1) * limit);

    const total = await User.countDocuments(query);

    res.json({
      success: true,
      data: users,
      pagination: {
        page: parseInt(page),
        limit: parseInt(limit),
        total,
        pages: Math.ceil(total / limit)
      }
    });

  } catch (error) {
    logger.error('Error in getAllUsers', {
      userId: req.user._id,
      error: error.message,
      stack: error.stack,
      operation: 'getAllUsers'
    });
    next(error);
  }
};

/**
 * @route   PUT /api/users/:userId/role
 * @desc    Update user role (Super Admin only)
 * @access  Private (Super Admin only)
 */
exports.updateUserRole = async (req, res, next) => {
  try {
    // Check if current user is super admin
    if (req.user.role !== 'super_admin') {
      return res.status(403).json({
        success: false,
        error: 'Access denied. Super admin privileges required.'
      });
    }

    const { userId } = req.params;
    const { role } = req.body;

    // Validate role
    const validRoles = ['super_admin', 'admin', 'agent', 'user'];
    if (!validRoles.includes(role)) {
      return res.status(400).json({
        success: false,
        error: 'Invalid role. Must be one of: super_admin, admin, agent, user'
      });
    }

    // Prevent demoting self
    if (userId === req.user._id.toString() && role !== 'super_admin') {
      return res.status(400).json({
        success: false,
        error: 'Cannot change your own super admin role'
      });
    }

    const user = await User.findByIdAndUpdate(
      userId,
      { role },
      { new: true, runValidators: true }
    ).select('-password_hash');

    if (!user) {
      return res.status(404).json({
        success: false,
        error: 'User not found'
      });
    }

    logger.business('User role updated successfully', {
      adminId: req.user._id,
      targetUserId: userId,
      oldRole: user.role,
      newRole: role,
      operation: 'updateUserRole'
    });

    res.json({
      success: true,
      message: 'User role updated successfully',
      data: user
    });

  } catch (error) {
    logger.error('Error in updateUserRole', {
      userId: req.user._id,
      targetUserId: req.params.userId,
      requestedRole: req.body.role,
      error: error.message,
      stack: error.stack,
      operation: 'updateUserRole'
    });
    next(error);
  }
};

/**
 * @route   PUT /api/users/:userId/credits
 * @desc    Update user wallet credits (Super Admin only)
 * @access  Private (Super Admin only)
 */
exports.updateUserCredits = async (req, res, next) => {
  try {
    // Check if current user is super admin
    if (req.user.role !== 'super_admin') {
      return res.status(403).json({
        success: false,
        error: 'Access denied. Super admin privileges required.'
      });
    }

    const { userId } = req.params;
    const { report_credits, enquiry_credits } = req.body;

    // Find user
    const user = await User.findById(userId);
    if (!user) {
      return res.status(404).json({
        success: false,
        error: 'User not found'
      });
    }

    // Find or create wallet
    let wallet = await Wallet.findOne({ user_id: userId });
    if (!wallet) {
      wallet = new Wallet({ user_id: userId });
    }

    // Update credits
    if (report_credits !== undefined) {
      wallet.report_credits = Math.max(0, report_credits);
    }
    if (enquiry_credits !== undefined) {
      wallet.enquiry_credits = Math.max(0, enquiry_credits);
    }

    await wallet.save();

    logger.business('User credits updated successfully', {
      adminId: req.user._id,
      targetUserId: userId,
      reportCredits: wallet.report_credits,
      enquiryCredits: wallet.enquiry_credits,
      operation: 'updateUserCredits'
    });

    res.json({
      success: true,
      message: 'User credits updated successfully',
      data: {
        user_id: userId,
        report_credits: wallet.report_credits,
        enquiry_credits: wallet.enquiry_credits
      }
    });

  } catch (error) {
    logger.error('Error in updateUserCredits', {
      userId: req.user._id,
      targetUserId: req.params.userId,
      requestedReportCredits: req.body.report_credits,
      requestedEnquiryCredits: req.body.enquiry_credits,
      error: error.message,
      stack: error.stack,
      operation: 'updateUserCredits'
    });
    next(error);
  }
};

/**
 * @route   DELETE /api/users/:userId
 * @desc    Delete user (Super Admin only)
 * @access  Private (Super Admin only)
 */
exports.deleteUser = async (req, res, next) => {
  try {
    // Check if current user is super admin
    if (req.user.role !== 'super_admin') {
      return res.status(403).json({
        success: false,
        error: 'Access denied. Super admin privileges required.'
      });
    }

    const { userId } = req.params;

    // Prevent self-deletion
    if (userId === req.user._id.toString()) {
      return res.status(400).json({
        success: false,
        error: 'Cannot delete your own account'
      });
    }

    // Find user
    const user = await User.findById(userId);
    if (!user) {
      return res.status(404).json({
        success: false,
        error: 'User not found'
      });
    }

    // Delete user and associated data
    await Wallet.findOneAndDelete({ user_id: userId });
    await User.findByIdAndDelete(userId);

    logger.business('User deleted successfully', {
      adminId: req.user._id,
      deletedUserId: userId,
      operation: 'deleteUser'
    });

    res.json({
      success: true,
      message: 'User deleted successfully'
    });

  } catch (error) {
    logger.error('Error in deleteUser', {
      userId: req.user._id,
      targetUserId: req.params.userId,
      error: error.message,
      stack: error.stack,
      operation: 'deleteUser'
    });
    next(error);
  }
};