const mongoose = require('mongoose');
const { Schema } = mongoose;

const withdrawalSchema = new Schema({
  agent_id: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  amount: { type: Number, required: true, min: 1 },
  status: { 
    type: String, 
    enum: ['pending', 'approved', 'rejected', 'processing', 'completed'], 
    default: 'pending' 
  },
  payment_method: { 
    type: String, 
    enum: ['bank_transfer', 'upi', 'phone_pe'], 
    required: true 
  },
  payment_details: {
    account_holder_name: { type: String },
    account_number: { type: String },
    ifsc_code: { type: String },
    bank_name: { type: String },
    upi_id: { type: String },
    phone_pe_number: { type: String }
  },
  admin_remarks: { type: String },
  processed_by: { type: Schema.Types.ObjectId, ref: 'User' },
  processed_at: { type: Date },
  transaction_id: { type: String } // For tracking payment
}, { timestamps: true });

// Index for faster queries
withdrawalSchema.index({ agent_id: 1, status: 1 });
withdrawalSchema.index({ status: 1, createdAt: -1 });

module.exports = mongoose.model('Withdrawal', withdrawalSchema);
