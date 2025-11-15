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
  status: { type: String }
}, { timestamps: true });

module.exports = mongoose.model('Report', reportSchema);