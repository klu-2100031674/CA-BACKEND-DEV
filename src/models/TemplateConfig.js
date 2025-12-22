const mongoose = require('mongoose');
const { Schema } = mongoose;

/**
 * Sheet Pricing Schema - For individual sheet pricing
 */
const sheetPricingSchema = new Schema({
  sheet_name: { type: String, required: true },
  display_name: { type: String },
  price: { type: Number, default: 0 },
  is_included: { type: Boolean, default: true }, // Whether this sheet is included in base price
  is_optional: { type: Boolean, default: false }, // Can user select/deselect this sheet
  is_visible: { type: Boolean, default: true }
}, { _id: false });

/**
 * Template Configuration Schema
 * Stores all template metadata, sheet configurations, and pricing
 */
const templateConfigSchema = new Schema({
  // Basic Info
  template_id: { 
    type: String, 
    required: true, 
    unique: true,
    index: true
  },
  name: { type: String, required: true },
  description: { type: String },
  version: { type: String, default: '1.0.0' },
  author: { type: String, default: 'CA' },
  
  // Template Type & Category
  report_type: { 
    type: String, 
    required: true,
    enum: ['CC', 'TERM_LOAN', 'HOUSING_LOAN', 'BUSINESS_LOAN', 'PERSONAL_LOAN', 'VEHICLE_LOAN', 'GOLD_LOAN', 'OTHER']
  },
  
  // Properties from meta.json
  properties: {
    no_of_years: { type: Number, default: 1 },
    type_of_report: { type: String }
  },
  
  // Sheet Configuration Arrays (from meta.json)
  initial_hide: [{ type: String }],
  initial_remove_formulas: [{ type: String }],
  after_generate_remove_formulas: [{ type: String }],
  after_generate_hide: [{ type: String }],
  after_generate_lock: [{ type: String }],
  
  // Default sheets to include in the full report (if not specified)
  full_report_sheets: [{ type: String }],
  
  // Analysis Sheets Configuration (for Term Loans)
  analysis_sheets: [{
    sheet_name: { type: String, required: true },
    display_name: { type: String },
    required: { type: Boolean, default: false },
    is_visible: { type: Boolean, default: true },
    price: { type: Number, default: 0 },
    amount_display: { type: String }
  }],
  
  // Pricing Configuration
  pricing: {
    base_price: { type: Number, default: 0 },
    credits_required: { type: Number, default: 1 },
    currency: { type: String, default: 'INR' },
    discount_percentage: { type: Number, default: 0 },
    // Per-sheet pricing
    sheet_pricing: [sheetPricingSchema]
  },
  
  // Excel File Reference
  excel_file: {
    filename: { type: String },
    path: { type: String },
    uploaded_at: { type: Date }
  },
  
  // Form Configuration (optional - for dynamic forms)
  form_config: {
    form_html_file: { type: String },
    form_fields: [{
      field_name: { type: String },
      field_type: { type: String },
      label: { type: String },
      required: { type: Boolean, default: false },
      default_value: { type: Schema.Types.Mixed },
      options: [{ type: String }], // For select/radio fields
      validation: { type: String }
    }]
  },
  
  // Status & Visibility
  is_active: { type: Boolean, default: true },
  is_featured: { type: Boolean, default: false },
  display_order: { type: Number, default: 0 },
  
  // Audit Fields
  created_by: { type: Schema.Types.ObjectId, ref: 'User' },
  updated_by: { type: Schema.Types.ObjectId, ref: 'User' }
}, { 
  timestamps: true,
  toJSON: { virtuals: true },
  toObject: { virtuals: true }
});

// Virtual for total price including all sheets
templateConfigSchema.virtual('total_price').get(function() {
  let total = this.pricing?.base_price || 0;
  if (this.pricing?.sheet_pricing) {
    total += this.pricing.sheet_pricing
      .filter(s => s.is_included && !s.is_optional)
      .reduce((sum, s) => sum + (s.price || 0), 0);
  }
  return total;
});

// Virtual for effective price after discount
templateConfigSchema.virtual('effective_price').get(function() {
  const total = this.total_price;
  const discount = this.pricing?.discount_percentage || 0;
  return total - (total * discount / 100);
});

// Index for searching
templateConfigSchema.index({ name: 'text', description: 'text' });
templateConfigSchema.index({ report_type: 1, is_active: 1 });

module.exports = mongoose.model('TemplateConfig', templateConfigSchema);
