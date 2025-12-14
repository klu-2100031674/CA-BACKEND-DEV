const mongoose = require('mongoose');
const { Schema } = mongoose;

const excelFileSchema = new Schema({
  user_id: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  template_id: { type: String, required: true },
  file_name: { type: String, required: true },
  file_data: { type: Buffer }, // Optional - files now stored in R2 cloud storage
  excel_r2_url: { type: String }, // R2 cloud storage URL for Excel file
  file_size: { type: Number, required: true },
  mime_type: { type: String, default: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  stage: { type: String, enum: ['stage1', 'stage2', 'final'], default: 'stage1' },
  generated_by: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  json_data: { type: Schema.Types.Mixed }, // Store the calculated JSON data
  all_sheets_data: { type: Schema.Types.Mixed }, // Store all sheets data
  formatted_wc_data: { type: Schema.Types.Mixed }, // Store formatted WC data
  html_content: { type: String }, // Store HTML content
  html_json_data: { type: Schema.Types.Mixed }, // Store HTML JSON data
  pdf_data: { type: Buffer }, // Optional - PDFs now stored in R2 cloud storage
  pdf_r2_url: { type: String }, // R2 cloud storage URL for PDF file
  pdf_file_name: { type: String },
  meta: { type: Schema.Types.Mixed }, // Store metadata (without local file paths)
  ai_report_generated: { type: Boolean, default: false },
  ai_report_data: { type: Buffer }, // Optional - AI reports stored in R2 cloud storage
  ai_report_r2_url: { type: String }, // R2 cloud storage URL for AI report PDF
  ai_report_file_name: { type: String },
  download_count: { type: Number, default: 0 },
  last_downloaded_at: { type: Date }
}, { timestamps: true });

// Index for efficient queries
excelFileSchema.index({ user_id: 1, createdAt: -1 });
excelFileSchema.index({ template_id: 1 });
excelFileSchema.index({ stage: 1 });

module.exports = mongoose.model('ExcelFile', excelFileSchema);