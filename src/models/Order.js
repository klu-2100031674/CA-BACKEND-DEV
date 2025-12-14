const mongoose = require('mongoose');
const { Schema } = mongoose;

const orderSchema = new Schema({
  user_id: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  
  // Pay-per-report model
  report_id: { type: Schema.Types.ObjectId, ref: 'Report' }, // Link to specific report
  template_id: { type: String }, // Template ID for the report
  report_title: { type: String }, // Report title for display
  
  // Legacy fields (deprecated but kept for backward compatibility)
  pack_type: { type: String, enum: ['report'], default: 'report' },
  credits: { type: Number, default: 1 }, // Always 1 for pay-per-report
  
  // Payment details
  amount_paid: { type: Number, required: true, min: 0 },
  currency: { type: String, default: 'INR' },
  status: { type: String, enum: ['paid', 'pending', 'failed'], default: 'pending' },
  payment_info: {
    rzp_order_id: String,
    rzp_payment_id: String,
    rzp_signature: String,
    method: String,
    captured: Boolean
  }
}, { timestamps: true });

module.exports = mongoose.model('Order', orderSchema);