/**
 * Template Service
 * Handles loading and caching of templates from database
 * Falls back to meta.json if database is empty
 */

const fs = require('fs').promises;
const path = require('path');
const TemplateConfig = require('../models/TemplateConfig');
const logger = require('../utils/logger');

// In-memory cache
let templatesCache = null;
let cacheTimestamp = null;
const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

/**
 * Clear the template cache
 */
function clearCache() {
  templatesCache = null;
  cacheTimestamp = null;
  logger.info('Template cache cleared');
}

/**
 * Load all active templates from database
 * Falls back to meta.json if database is empty
 */
async function loadTemplates(forceRefresh = false) {
  // Return cached data if valid
  if (!forceRefresh && templatesCache && cacheTimestamp && (Date.now() - cacheTimestamp < CACHE_DURATION)) {
    return templatesCache;
  }
  
  try {
    // Try loading from database first
    const dbTemplates = await TemplateConfig.find({ is_active: true })
      .sort({ display_order: 1, createdAt: -1 });
    
    if (dbTemplates && dbTemplates.length > 0) {
      // Convert to meta.json format for backward compatibility
      templatesCache = dbTemplates.map(convertToMetaFormat);
      cacheTimestamp = Date.now();
      logger.info('Templates loaded from database', { templateCount: templatesCache.length });
      return templatesCache;
    }
    
    // Fallback to meta.json if database is empty
    logger.warn('No templates in database, falling back to meta.json');
    return await loadFromMetaJson();
    
  } catch (error) {
    logger.error('Failed to load templates from database', { error: error.message });
    // Fallback to meta.json on database error
    return await loadFromMetaJson();
  }
}

/**
 * Load templates from meta.json (fallback)
 */
async function loadFromMetaJson() {
  try {
    const metaPath = path.join(__dirname, '../../templates/meta.json');
    logger.debug('Loading templates from meta.json', { metaPath });
    const metaData = await fs.readFile(metaPath, 'utf8');
    templatesCache = JSON.parse(metaData);
    cacheTimestamp = Date.now();
    logger.info('Templates loaded from meta.json', { templateCount: templatesCache.length });
    return templatesCache;
  } catch (error) {
    logger.error('Failed to load templates from meta.json', { error: error.message });
    throw new Error('Failed to load templates: ' + error.message);
  }
}

/**
 * Get a single template by ID from database or cache
 */
async function getTemplateById(templateId) {
  try {
    // Try database first
    const dbTemplate = await TemplateConfig.findOne({ 
      template_id: templateId,
      is_active: true 
    });
    
    if (dbTemplate) {
      return convertToMetaFormat(dbTemplate);
    }
    
    // Fallback to cache/meta.json
    const templates = await loadTemplates();
    return templates.find(t => t.id === templateId || t.id.toLowerCase() === templateId.toLowerCase());
    
  } catch (error) {
    logger.error('Failed to get template by ID', { templateId, error: error.message });
    // Fallback to cache
    const templates = await loadTemplates();
    return templates.find(t => t.id === templateId || t.id.toLowerCase() === templateId.toLowerCase());
  }
}

/**
 * Get template config from database (full model)
 */
async function getTemplateConfig(templateId) {
  const template = await TemplateConfig.findOne({ 
    template_id: templateId
  });
  return template;
}

/**
 * Convert database template to meta.json format for backward compatibility
 */
function convertToMetaFormat(dbTemplate) {
  return {
    id: dbTemplate.template_id,
    name: dbTemplate.name,
    description: dbTemplate.description || '',
    version: dbTemplate.version || '1.0.0',
    author: dbTemplate.author || 'CA',
    initialHide: dbTemplate.initial_hide || [],
    initialRemoveFormulas: dbTemplate.initial_remove_formulas || [],
    afterGenerateRemoveFormulas: dbTemplate.after_generate_remove_formulas || [],
    afterGenerateHide: dbTemplate.after_generate_hide || [],
    afterGenerateLock: dbTemplate.after_generate_lock || [],
    lastModified: dbTemplate.updatedAt?.toISOString() || new Date().toISOString(),
    createdAt: dbTemplate.createdAt?.toISOString() || new Date().toISOString(),
    properties: {
      'No Of Years': dbTemplate.properties?.no_of_years || 1,
      'Type of Report': dbTemplate.properties?.type_of_report || dbTemplate.report_type || 'CC'
    },
    // Additional fields from database
    _id: dbTemplate._id,
    pricing: dbTemplate.pricing,
    is_active: dbTemplate.is_active,
    is_featured: dbTemplate.is_featured,
    display_order: dbTemplate.display_order,
    report_type: dbTemplate.report_type
  };
}

/**
 * Convert meta.json format to database format
 */
function convertToDbFormat(metaTemplate) {
  return {
    template_id: metaTemplate.id,
    name: metaTemplate.name,
    description: metaTemplate.description || '',
    version: metaTemplate.version || '1.0.0',
    author: metaTemplate.author || 'CA',
    report_type: mapReportType(metaTemplate.properties?.['Type of Report']),
    properties: {
      no_of_years: metaTemplate.properties?.['No Of Years'] || 1,
      type_of_report: metaTemplate.properties?.['Type of Report'] || 'CC'
    },
    initial_hide: metaTemplate.initialHide || [],
    initial_remove_formulas: metaTemplate.initialRemoveFormulas || [],
    after_generate_remove_formulas: metaTemplate.afterGenerateRemoveFormulas || [],
    after_generate_hide: metaTemplate.afterGenerateHide || [],
    after_generate_lock: metaTemplate.afterGenerateLock || []
  };
}

/**
 * Map report type string to enum value
 */
function mapReportType(type) {
  const typeMap = {
    'CC': 'CC',
    'TERM_LOAN': 'TERM_LOAN',
    'Term Loan': 'TERM_LOAN',
    'HOUSING_LOAN': 'HOUSING_LOAN',
    'BUSINESS_LOAN': 'BUSINESS_LOAN',
    'PERSONAL_LOAN': 'PERSONAL_LOAN',
    'VEHICLE_LOAN': 'VEHICLE_LOAN',
    'GOLD_LOAN': 'GOLD_LOAN'
  };
  return typeMap[type] || 'OTHER';
}

/**
 * Update template in database and clear cache
 */
async function updateTemplate(templateId, updateData) {
  const template = await TemplateConfig.findOneAndUpdate(
    { template_id: templateId },
    updateData,
    { new: true }
  );
  
  // Clear cache after update
  clearCache();
  
  return template;
}

/**
 * Get templates with pricing information
 */
async function getTemplatesWithPricing() {
  const templates = await TemplateConfig.find({ is_active: true })
    .select('template_id name description report_type pricing is_featured display_order')
    .sort({ display_order: 1, createdAt: -1 });
  
  return templates;
}

module.exports = {
  loadTemplates,
  loadFromMetaJson,
  getTemplateById,
  getTemplateConfig,
  convertToMetaFormat,
  convertToDbFormat,
  updateTemplate,
  getTemplatesWithPricing,
  clearCache
};
