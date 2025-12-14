const express = require('express');
const TemplateConfig = require('../models/TemplateConfig');
const { verifyToken, requireRole } = require('../middleware/auth');
const templateService = require('../services/templateService');
const router = express.Router();
const fs = require('fs');
const path = require('path');

/**
 * Template Configuration Routes
 * Manages Excel template configurations and pricing
 */

// ============================================================================
// Public Routes
// ============================================================================

// Get all active templates (public)
router.get('/', async (req, res) => {
  try {
    const { report_type, is_featured, search } = req.query;
    
    const query = { is_active: true };
    
    if (report_type) {
      query.report_type = report_type.toUpperCase();
    }
    
    if (is_featured === 'true') {
      query.is_featured = true;
    }
    
    if (search) {
      query.$text = { $search: search };
    }
    
    const templates = await TemplateConfig.find(query)
      .select('-form_config.form_fields -created_by -updated_by')
      .sort({ display_order: 1, createdAt: -1 });
    
    res.json({
      success: true,
      data: templates
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get template by ID (public)
router.get('/:templateId', async (req, res) => {
  try {
    const { templateId } = req.params;
    
    const template = await TemplateConfig.findOne({ 
      template_id: templateId,
      is_active: true 
    });
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    res.json({
      success: true,
      data: template
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get template pricing (public)
router.get('/:templateId/pricing', async (req, res) => {
  try {
    const { templateId } = req.params;
    
    const template = await TemplateConfig.findOne({ 
      template_id: templateId,
      is_active: true 
    }).select('template_id name pricing');
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    res.json({
      success: true,
      data: {
        template_id: template.template_id,
        name: template.name,
        pricing: template.pricing,
        total_price: template.total_price,
        effective_price: template.effective_price
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ============================================================================
// Admin Routes
// ============================================================================

// Get all templates (admin - includes inactive)
router.get('/admin/all', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { report_type, is_active, page = 1, limit = 20 } = req.query;
    
    const query = {};
    
    if (report_type) {
      query.report_type = report_type.toUpperCase();
    }
    
    if (is_active !== undefined) {
      query.is_active = is_active === 'true';
    }
    
    const skip = (parseInt(page) - 1) * parseInt(limit);
    
    const [templates, total] = await Promise.all([
      TemplateConfig.find(query)
        .populate('created_by', 'name email')
        .populate('updated_by', 'name email')
        .sort({ display_order: 1, createdAt: -1 })
        .skip(skip)
        .limit(parseInt(limit)),
      TemplateConfig.countDocuments(query)
    ]);
    
    res.json({
      success: true,
      data: {
        templates,
        pagination: {
          page: parseInt(page),
          limit: parseInt(limit),
          total,
          pages: Math.ceil(total / parseInt(limit))
        }
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Create new template config
router.post('/', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const {
      template_id,
      name,
      description,
      version,
      author,
      report_type,
      properties,
      initial_hide,
      initial_remove_formulas,
      after_generate_remove_formulas,
      after_generate_hide,
      after_generate_lock,
      pricing,
      excel_file,
      form_config,
      is_active,
      is_featured,
      display_order
    } = req.body;
    
    // Check if template_id already exists
    const existing = await TemplateConfig.findOne({ template_id });
    if (existing) {
      return res.status(400).json({ error: 'Template ID already exists' });
    }
    
    const template = new TemplateConfig({
      template_id,
      name,
      description,
      version,
      author,
      report_type: report_type?.toUpperCase() || 'OTHER',
      properties,
      initial_hide: initial_hide || [],
      initial_remove_formulas: initial_remove_formulas || [],
      after_generate_remove_formulas: after_generate_remove_formulas || [],
      after_generate_hide: after_generate_hide || [],
      after_generate_lock: after_generate_lock || [],
      pricing: pricing || { base_price: 0, credits_required: 1 },
      excel_file,
      form_config,
      is_active: is_active !== undefined ? is_active : true,
      is_featured: is_featured || false,
      display_order: display_order || 0,
      created_by: req.user._id,
      updated_by: req.user._id
    });
    
    await template.save();
    
    // Clear template cache
    templateService.clearCache();
    
    res.status(201).json({
      success: true,
      message: 'Template configuration created successfully',
      data: template
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Update template config
router.put('/:templateId', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { templateId } = req.params;
    const updateData = req.body;
    
    // Find by template_id (string) or MongoDB _id
    let template = await TemplateConfig.findOne({ template_id: templateId });
    if (!template) {
      template = await TemplateConfig.findById(templateId);
    }
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // Update allowed fields
    const allowedFields = [
      'name', 'description', 'version', 'author', 'report_type', 'properties',
      'initial_hide', 'initial_remove_formulas', 'after_generate_remove_formulas',
      'after_generate_hide', 'after_generate_lock', 'pricing', 'excel_file',
      'form_config', 'is_active', 'is_featured', 'display_order'
    ];
    
    allowedFields.forEach(field => {
      if (updateData[field] !== undefined) {
        template[field] = updateData[field];
      }
    });
    
    template.updated_by = req.user._id;
    await template.save();
    
    // Clear template cache
    templateService.clearCache();
    
    res.json({
      success: true,
      message: 'Template configuration updated successfully',
      data: template
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Update template pricing only
router.patch('/:templateId/pricing', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { templateId } = req.params;
    const { base_price, credits_required, currency, discount_percentage, sheet_pricing } = req.body;
    
    let template = await TemplateConfig.findOne({ template_id: templateId });
    if (!template) {
      template = await TemplateConfig.findById(templateId);
    }
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // Update pricing fields
    if (base_price !== undefined) template.pricing.base_price = base_price;
    if (credits_required !== undefined) template.pricing.credits_required = credits_required;
    if (currency !== undefined) template.pricing.currency = currency;
    if (discount_percentage !== undefined) template.pricing.discount_percentage = discount_percentage;
    if (sheet_pricing !== undefined) template.pricing.sheet_pricing = sheet_pricing;
    
    template.updated_by = req.user._id;
    await template.save();
    
    // Clear template cache
    templateService.clearCache();
    
    res.json({
      success: true,
      message: 'Template pricing updated successfully',
      data: {
        template_id: template.template_id,
        name: template.name,
        pricing: template.pricing,
        total_price: template.total_price,
        effective_price: template.effective_price
      }
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Toggle template status
router.patch('/:templateId/status', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { templateId } = req.params;
    const { is_active } = req.body;
    
    let template = await TemplateConfig.findOne({ template_id: templateId });
    if (!template) {
      template = await TemplateConfig.findById(templateId);
    }
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    template.is_active = is_active;
    template.updated_by = req.user._id;
    await template.save();
    
    // Clear template cache
    templateService.clearCache();
    
    res.json({
      success: true,
      message: `Template ${is_active ? 'activated' : 'deactivated'} successfully`,
      data: template
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Delete template
router.delete('/:templateId', verifyToken, requireRole(['super_admin']), async (req, res) => {
  try {
    const { templateId } = req.params;
    
    let template = await TemplateConfig.findOne({ template_id: templateId });
    if (!template) {
      template = await TemplateConfig.findById(templateId);
    }
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    await TemplateConfig.deleteOne({ _id: template._id });
    
    // Clear template cache
    templateService.clearCache();
    
    res.json({
      success: true,
      message: 'Template deleted successfully'
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// ============================================================================
// Migration Route - Import from meta.json
// ============================================================================

router.post('/migrate-from-meta', verifyToken, requireRole(['super_admin']), async (req, res) => {
  try {
    const metaPath = path.join(__dirname, '../../templates/meta.json');
    
    if (!fs.existsSync(metaPath)) {
      return res.status(404).json({ error: 'meta.json not found' });
    }
    
    const metaData = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
    
    const results = {
      created: 0,
      updated: 0,
      errors: []
    };
    
    for (const item of metaData) {
      try {
        const templateData = {
          template_id: item.id,
          name: item.name,
          description: item.description,
          version: item.version || '1.0.0',
          author: item.author || 'CA',
          report_type: item.properties?.['Type of Report']?.toUpperCase() || 'OTHER',
          properties: {
            no_of_years: item.properties?.['No Of Years'] || 1,
            type_of_report: item.properties?.['Type of Report']
          },
          initial_hide: item.initialHide || [],
          initial_remove_formulas: item.initialRemoveFormulas || [],
          after_generate_remove_formulas: item.afterGenerateRemoveFormulas || [],
          after_generate_hide: item.afterGenerateHide || [],
          after_generate_lock: item.afterGenerateLock || [],
          pricing: {
            base_price: 500, // Default price
            credits_required: 1,
            currency: 'INR',
            sheet_pricing: []
          },
          is_active: true,
          created_by: req.user._id,
          updated_by: req.user._id
        };
        
        // Check if exists
        const existing = await TemplateConfig.findOne({ template_id: item.id });
        
        if (existing) {
          // Update existing
          Object.assign(existing, templateData);
          await existing.save();
          results.updated++;
        } else {
          // Create new
          const template = new TemplateConfig(templateData);
          await template.save();
          results.created++;
        }
      } catch (err) {
        results.errors.push({ id: item.id, error: err.message });
      }
    }
    
    // Clear template cache after migration
    templateService.clearCache();
    
    res.json({
      success: true,
      message: 'Migration completed',
      data: results
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get report types for dropdown
router.get('/meta/report-types', async (req, res) => {
  res.json({
    success: true,
    data: [
      { value: 'CC', label: 'Cash Credit (CC)' },
      { value: 'TERM_LOAN', label: 'Term Loan' },
      { value: 'HOUSING_LOAN', label: 'Housing Loan' },
      { value: 'BUSINESS_LOAN', label: 'Business Loan' },
      { value: 'PERSONAL_LOAN', label: 'Personal Loan' },
      { value: 'VEHICLE_LOAN', label: 'Vehicle Loan' },
      { value: 'GOLD_LOAN', label: 'Gold Loan' },
      { value: 'OTHER', label: 'Other' }
    ]
  });
});

module.exports = router;
