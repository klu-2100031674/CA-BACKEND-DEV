const mongoose = require('mongoose');
const { Schema } = mongoose;

const commissionSchema = new Schema({
  agent_id: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  user_id: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  order_id: { type: Schema.Types.ObjectId, ref: 'Order', required: true },
  rate_percent: { type: Number, required: true, min: 0, max: 100 },
  base_amount: { type: Number, required: true, min: 0 },
  commission_amount: { type: Number, required: true, min: 0 },
  status: { type: String, enum: ['accrued', 'paid'], default: 'accrued' }
}, { timestamps: true });

module.exports = mongoose.model('Commission', commissionSchema);