/**
 * Cleanup script to delete old reports without R2 cloud storage URLs
 * Run this once to clean up reports generated before cloud migration
 */

const mongoose = require('mongoose');
const Report = require('./src/models/Report');
require('dotenv').config();

async function cleanupOldReports() {
  try {
    await mongoose.connect(process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/ca_report_generation');
    console.log('✅ Connected to MongoDB');

    // Find reports without R2 URLs (old reports with local paths or no files)
    const oldReports = await Report.find({
      $or: [
        { pdf_file_url: { $exists: false } },
        { pdf_file_url: null },
        { pdf_file_url: { $not: /r2\.cloudflarestorage\.com/ } }
      ]
    }).select('_id title createdAt pdf_file_url validation_status');

    console.log(`\n📊 Found ${oldReports.length} old reports without R2 cloud storage\n`);

    if (oldReports.length === 0) {
      console.log('✨ No old reports to clean up!');
      process.exit(0);
      return;
    }

    // Show details
    oldReports.forEach((report, index) => {
      console.log(`${index + 1}. ${report.title || 'Untitled'}`);
      console.log(`   ID: ${report._id}`);
      console.log(`   Created: ${report.createdAt}`);
      console.log(`   Status: ${report.validation_status}`);
      console.log(`   PDF URL: ${report.pdf_file_url || 'none'}`);
      console.log('');
    });

    // Delete old reports
    const result = await Report.deleteMany({
      $or: [
        { pdf_file_url: { $exists: false } },
        { pdf_file_url: null },
        { pdf_file_url: { $not: /r2\.cloudflarestorage\.com/ } }
      ]
    });

    console.log(`\n🗑️  Deleted ${result.deletedCount} old reports`);
    console.log('✅ Cleanup complete! All remaining reports use R2 cloud storage.\n');

  } catch (error) {
    console.error('❌ Error:', error.message);
  } finally {
    await mongoose.connection.close();
    process.exit(0);
  }
}

cleanupOldReports();
