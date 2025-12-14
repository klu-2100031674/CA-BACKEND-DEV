/**
 * Migration Script: Import templates from meta.json to MongoDB
 * Run this script once to migrate existing templates to database
 * 
 * Usage: node src/scripts/migrateTemplates.js
 */

require('dotenv').config();
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');
const TemplateConfig = require('../models/TemplateConfig');
const config = require('../config/environment');

async function migrateTemplates() {
  try {
    console.log('🚀 Starting template migration...');
    
    // Connect to MongoDB
    await mongoose.connect(config.MONGODB_URI);
    console.log('✅ Connected to MongoDB');
    
    // Load meta.json
    const metaPath = path.join(__dirname, '../../templates/meta.json');
    
    if (!fs.existsSync(metaPath)) {
      console.error('❌ meta.json not found at:', metaPath);
      process.exit(1);
    }
    
    const metaData = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
    console.log(`📄 Found ${metaData.length} templates in meta.json`);
    
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
          description: item.description || '',
          version: item.version || '1.0.0',
          author: item.author || 'CA',
          report_type: mapReportType(item.properties?.['Type of Report']),
          properties: {
            no_of_years: item.properties?.['No Of Years'] || 1,
            type_of_report: item.properties?.['Type of Report'] || 'CC'
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
            discount_percentage: 0,
            sheet_pricing: []
          },
          excel_file: {
            filename: `${item.id}.xlsx`,
            path: `templates/excel/${item.id}.xlsx`
          },
          is_active: true,
          is_featured: false,
          display_order: 0
        };
        
        // Check if template already exists
        const existing = await TemplateConfig.findOne({ template_id: item.id });
        
        if (existing) {
          // Update existing template
          Object.assign(existing, templateData);
          await existing.save();
          console.log(`📝 Updated: ${item.name}`);
          results.updated++;
        } else {
          // Create new template
          const template = new TemplateConfig(templateData);
          await template.save();
          console.log(`✨ Created: ${item.name}`);
          results.created++;
        }
      } catch (err) {
        console.error(`❌ Error processing ${item.id}:`, err.message);
        results.errors.push({ id: item.id, error: err.message });
      }
    }
    
    console.log('\n📊 Migration Summary:');
    console.log(`   Created: ${results.created}`);
    console.log(`   Updated: ${results.updated}`);
    console.log(`   Errors: ${results.errors.length}`);
    
    if (results.errors.length > 0) {
      console.log('\n❌ Errors:');
      results.errors.forEach(e => console.log(`   - ${e.id}: ${e.error}`));
    }
    
    console.log('\n✅ Migration completed!');
    
  } catch (error) {
    console.error('❌ Migration failed:', error);
  } finally {
    await mongoose.disconnect();
    console.log('🔌 Disconnected from MongoDB');
  }
}

function mapReportType(type) {
  const typeMap = {
    'CC': 'CC',
    'TERM_LOAN': 'TERM_LOAN',
    'Term Loan': 'TERM_LOAN',
    'HOUSING_LOAN': 'HOUSING_LOAN',
    'Housing Loan': 'HOUSING_LOAN',
    'BUSINESS_LOAN': 'BUSINESS_LOAN',
    'Business Loan': 'BUSINESS_LOAN',
    'PERSONAL_LOAN': 'PERSONAL_LOAN',
    'Personal Loan': 'PERSONAL_LOAN',
    'VEHICLE_LOAN': 'VEHICLE_LOAN',
    'Vehicle Loan': 'VEHICLE_LOAN',
    'GOLD_LOAN': 'GOLD_LOAN',
    'Gold Loan': 'GOLD_LOAN'
  };
  
  return typeMap[type] || 'OTHER';
}

// Run migration
migrateTemplates();
