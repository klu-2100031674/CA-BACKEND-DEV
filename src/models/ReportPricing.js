const mongoose = require('mongoose');
const { Schema } = mongoose;

const reportPricingSchema = new Schema({
  report_type: { 
    type: String, 
    required: true, 
    unique: true 
  }, // e.g., 'CC', 'TERM_LOAN', etc.
  name: { type: String, required: true }, // Display name
  description: { type: String },
  price_per_credit: { type: Number, required: true, min: 0 }, // Price in INR
  credits_required: { type: Number, required: true, default: 1, min: 1 },
  is_active: { type: Boolean, default: true },
  created_by: { type: Schema.Types.ObjectId, ref: 'User' },
  updated_by: { type: Schema.Types.ObjectId, ref: 'User' }
}, { timestamps: true });

module.exports = mongoose.model('ReportPricing', reportPricingSchema);
