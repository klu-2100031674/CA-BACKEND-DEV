const mongoose = require('mongoose');
const { Schema } = mongoose;

const bankDetailsSchema = new Schema({
  account_holder_name: { type: String },
  account_number: { type: String },
  ifsc_code: { type: String },
  bank_name: { type: String },
  branch: { type: String },
  upi_id: { type: String },
  phone_pe_number: { type: String }
}, { _id: false });

const userSchema = new Schema({
  role: { type: String, enum: ['super_admin', 'admin', 'agent', 'user'], required: true, default: 'user' },
  name: { type: String, required: true },
  email: { type: String, required: true, unique: true },
  email_verified: { type: Boolean, default: false },
  password_hash: { type: String, required: true },
  agent_id: { type: Schema.Types.ObjectId, ref: 'User', default: null }, // Referral - which agent referred this user
  profile_logo: { type: String }, // URL or path to profile logo
  company_logo: { type: String }, // URL or path to company logo
  signature_url: { type: String }, // URL or path to admin signature image
  company_name: { type: String },
  phone: { type: String },
  address: { type: String },
  is_active: { type: Boolean, default: true }, // Enable/disable user
  
  // Agent-specific fields
  referral_code: { type: String, unique: true, sparse: true }, // Unique referral code for agents
  commission_rate: { type: Number, default: 10, min: 0, max: 100 }, // Commission percentage for agents
  bank_details: { type: bankDetailsSchema }, // Bank details for withdrawal
  
  // Password reset OTP
  reset_otp: { type: String },
  reset_otp_expires: { type: Date }
}, { timestamps: true });

// Generate referral code for agents
userSchema.pre('save', function(next) {
  if (this.role === 'agent' && !this.referral_code) {
    // Generate a unique referral code
    this.referral_code = `AG${this.name.substring(0, 3).toUpperCase()}${Date.now().toString(36).toUpperCase()}`;
  }
  next();
});

// Virtual for referred users count
userSchema.virtual('referred_users', {
  ref: 'User',
  localField: '_id',
  foreignField: 'agent_id',
  count: true
});

module.exports = mongoose.model('User', userSchema);
