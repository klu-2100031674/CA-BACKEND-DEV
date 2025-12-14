const express = require('express');
const ReportPricing = require('../models/ReportPricing');
const { verifyToken, requireRole } = require('../middleware/auth');
const router = express.Router();

// Get all report pricing (public for display, or all for admin)
router.get('/', async (req, res) => {
  try {
    const query = req.user?.role === 'super_admin' || req.user?.role === 'admin' 
      ? {} 
      : { is_active: true };
    
    const pricing = await ReportPricing.find(query)
      .populate('created_by', 'name email')
      .populate('updated_by', 'name email')
      .sort({ report_type: 1 });
    
    res.json({
      success: true,
      data: pricing
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get pricing for a specific report type (reads from Template model)
router.get('/:reportType', async (req, res) => {
  try {
    const { reportType } = req.params;
    const lookupId = reportType.toUpperCase();
    
    console.log(`🔍 [getPricing] Looking up template pricing for: "${lookupId}"`);
    
    const TemplateConfig = require('../models/TemplateConfig');
    const template = await TemplateConfig.findOne({ template_id: lookupId, is_active: true });
    
    if (!template || !template.pricing) {
      console.log(`⚠️ [getPricing] No template pricing found, returning default`);
      return res.json({
        success: true,
        data: {
          report_type: lookupId,
          name: reportType,
          price_per_credit: 100,
          credits_required: 1,
          is_default: true
        }
      });
    }
    
    // Return pricing from template
    const amount = template.pricing.effective_price || template.pricing.total_price || template.pricing.base_price || 100;
    
    console.log(`✅ [getPricing] Found template pricing: ₹${amount}`, {
      hasSheetPricing: !!template.pricing.sheet_pricing,
      sheetCount: template.pricing.sheet_pricing?.length || 0
    });
    
    res.json({
      success: true,
      data: {
        report_type: lookupId,
        name: template.name,
        price_per_credit: amount,
        credits_required: template.pricing.credits_required || 1,
        base_price: template.pricing.base_price,
        total_price: template.pricing.total_price,
        effective_price: template.pricing.effective_price,
        currency: template.pricing.currency || 'INR',
        discount_percentage: template.pricing.discount_percentage || 0,
        sheet_pricing: template.pricing.sheet_pricing || [],
        is_default: false
      }
    });
  } catch (error) {
    console.error(`❌ [getPricing] Error:`, error);
    res.status(500).json({ error: error.message });
  }
});

// Admin: Create or update pricing
router.post('/', verifyToken, requireRole(['super_admin', 'admin']), async (req, res) => {
  try {
    const { report_type, name, description, price_per_credit, credits_required, is_active } = req.body;
    
    if (!report_type || !name || price_per_credit === undefined) {
      return res.status(400).json({ error: 'report_type, name, and price_per_credit are required' });
    }
    
    const existingPricing = await ReportPricing.findOne({ report_type: report_type.toUpperCase() });
    
    if (existingPricing) {
      // Update existing
      existingPricing.name = name;
      existingPricing.description = description;
      existingPricing.price_per_credit = price_per_credit;
      existingPricing.credits_required = credits_required || 1;
      existingPricing.is_active = is_active !== undefined ? is_active : true;
      existingPricing.updated_by = req.user._id;
      
      await existingPricing.save();
      
      res.json({
        success: true,
        message: 'Pricing updated successfully',
        data: existingPricing
      });
    } else {
      // Create new
      const pricing = new ReportPricing({
        report_type: report_type.toUpperCase(),
        name,
        description,
        price_per_credit,
        credits_required: credits_required || 1,
        is_active: is_active !== undefined ? is_active : true,
        created_by: req.user._id,
        updated_by: req.user._id
      });
      
      await pricing.save();
      
      res.status(201).json({
        success: true,
        message: 'Pricing created successfully',
        data: pricing
      });
    }
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Admin: Update pricing
router.put('/:id', verifyToken, requireRole(['super_admin', 'admin']), async (req, res) => {
  try {
    const { id } = req.params;
    const { name, description, price_per_credit, credits_required, is_active } = req.body;
    
    const pricing = await ReportPricing.findById(id);
    
    if (!pricing) {
      return res.status(404).json({ error: 'Pricing not found' });
    }
    
    if (name) pricing.name = name;
    if (description !== undefined) pricing.description = description;
    if (price_per_credit !== undefined) pricing.price_per_credit = price_per_credit;
    if (credits_required !== undefined) pricing.credits_required = credits_required;
    if (is_active !== undefined) pricing.is_active = is_active;
    pricing.updated_by = req.user._id;
    
    await pricing.save();
    
    res.json({
      success: true,
      message: 'Pricing updated successfully',
      data: pricing
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Admin: Delete pricing
router.delete('/:id', verifyToken, requireRole(['super_admin', 'admin']), async (req, res) => {
  try {
    const { id } = req.params;
    
    const pricing = await ReportPricing.findByIdAndDelete(id);
    
    if (!pricing) {
      return res.status(404).json({ error: 'Pricing not found' });
    }
    
    res.json({
      success: true,
      message: 'Pricing deleted successfully'
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

module.exports = router;
