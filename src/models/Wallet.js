const mongoose = require('mongoose');
const { Schema } = mongoose;

const walletSchema = new Schema({
  user_id: { type: Schema.Types.ObjectId, ref: 'User', required: true, unique: true },
  report_credits: { type: Number, default: 0, min: 0 },
  enquiry_credits: { type: Number, default: 0, min: 0 },
  notes: { type: String }
}, { timestamps: true });

module.exports = mongoose.model('Wallet', walletSchema);