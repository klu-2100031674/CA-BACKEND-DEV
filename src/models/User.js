const mongoose = require('mongoose');
const { Schema } = mongoose;

const userSchema = new Schema({
  role: { type: String, enum: ['super_admin', 'admin', 'agent', 'user'], required: true, default: 'user' },
  name: { type: String, required: true },
  email: { type: String, required: true, unique: true },
  email_verified: { type: Boolean, default: false },
  password_hash: { type: String, required: true },
  agent_id: { type: Schema.Types.ObjectId, ref: 'User', default: null },
  profile_logo: { type: String }, // URL or path to profile logo
  company_logo: { type: String }, // URL or path to company logo
  company_name: { type: String },
  phone: { type: String },
  address: { type: String }
}, { timestamps: true });

module.exports = mongoose.model('User', userSchema);
