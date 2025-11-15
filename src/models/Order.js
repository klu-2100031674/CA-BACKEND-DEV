const mongoose = require('mongoose');
const { Schema } = mongoose;

const orderSchema = new Schema({
  user_id: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  pack_type: { type: String, enum: ['report'], required: true },
  credits: { type: Number, required: true, min: 1 },
  amount_paid: { type: Number, required: true, min: 0 },
  currency: { type: String, default: 'INR' },
  status: { type: String, enum: ['paid', 'pending', 'failed'], default: 'pending' },
  payment_info: {
    rzp_order_id: String,
    rzp_payment_id: String,
    method: String,
    captured: Boolean
  }
}, { timestamps: true });

module.exports = mongoose.model('Order', orderSchema);