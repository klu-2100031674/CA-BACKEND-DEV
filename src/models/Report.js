const mongoose = require('mongoose');
const { Schema } = mongoose;

const reportSchema = new Schema({
  user_id: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  title: { type: String, required: true },
  templateId : { type: String, required: true },
  excel_file_url: { type: String },
  excel_data: { type: Buffer }, // Store Excel file data directly in DB
  json_data: { type: Schema.Types.Mixed }, // Store JSON data for browser display
  pdf_file_url: { type: String },
  json_file_url: { type: String }, // Points to {reportId}.json file with final Excel data
  compressed_json_url: { type: String }, // Points to {reportId}.json.gz file with final Excel data (compressed)
  hidden_sheets : [String], // Stores afterGenerateHide from meta.json
  locked_sheets : [String], // Stores afterGenerateLock from meta.json
  form_data: { type: Schema.Types.Mixed }, // Store user input data from generate page (small payload)
  report_metadata: { type: Schema.Types.Mixed }, // Optional small metadata (not the full Excel JSON)
  status: { type: String, default: 'draft' },
  
  // Payment Information (Pay-per-report model)
  payment: {
    status: { type: String, enum: ['pending', 'completed', 'failed', 'refunded'], default: 'pending' },
    amount: { type: Number, default: 0 },
    currency: { type: String, default: 'INR' },
    razorpay_order_id: { type: String }, // Razorpay order ID
    razorpay_payment_id: { type: String }, // Razorpay payment ID after success
    razorpay_signature: { type: String }, // Razorpay signature for verification
    transaction_id: { type: String }, // Legacy field
    payment_method: { type: String },
    paid_at: { type: Date }
  },
  
  // Validation/Approval Workflow
  validation_status: { 
    type: String, 
    enum: ['draft', 'pending_payment', 'pending_validation', 'under_review', 'approved', 'rejected'],
    default: 'draft'
  },
  validated_by: { type: Schema.Types.ObjectId, ref: 'User' },
  validated_at: { type: Date },
  validation_notes: { type: String }, // Admin notes (internal)
  rejection_reason: { type: String }, // Reason shown to user if rejected
  approval_email_sent: { type: Boolean, default: false },
  
  // Report Type for better categorization
  report_type: { type: String }, // CC, TERM_LOAN, etc.
  
  // Client Information (for display in admin)
  client_name: { type: String },
  client_details: { type: Schema.Types.Mixed }
}, { timestamps: true });

// Index for efficient queries
reportSchema.index({ validation_status: 1, createdAt: -1 });
reportSchema.index({ user_id: 1, validation_status: 1 });

module.exports = mongoose.model('Report', reportSchema);