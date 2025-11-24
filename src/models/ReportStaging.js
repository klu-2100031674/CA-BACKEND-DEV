const mongoose = require('mongoose');
const { Schema } = mongoose;

const reportStagingSchema = new Schema({
  user_id: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  template_id: { type: String, required: true },
  excel_data: { type: Buffer, required: true }, // Store Excel file data directly in DB
  file_name: { type: String, required: true },
  status: { type: String, default: 'active' } // active, used, expired
}, { timestamps: true });

// Index for efficient queries
reportStagingSchema.index({ user_id: 1, template_id: 1, createdAt: -1 });

module.exports = mongoose.model('ReportStaging', reportStagingSchema);