/**
 * Archive old reports without R2 cloud storage URLs
 * Marks them as 'archived' instead of deleting
 */

const mongoose = require('mongoose');
const Report = require('./src/models/Report');
require('dotenv').config();

async function archiveOldReports() {
  try {
    await mongoose.connect(process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/ca_report_generation');
    console.log('✅ Connected to MongoDB');

    // Find reports without R2 URLs
    const oldReports = await Report.find({
      $or: [
        { pdf_file_url: { $exists: false } },
        { pdf_file_url: null },
        { pdf_file_url: { $not: /r2\.cloudflarestorage\.com/ } }
      ]
    }).select('_id title createdAt pdf_file_url validation_status');

    console.log(`\n📊 Found ${oldReports.length} old reports to archive\n`);

    if (oldReports.length === 0) {
      console.log('✨ No old reports to archive!');
      process.exit(0);
      return;
    }

    // Update reports to archived status
    const result = await Report.updateMany(
      {
        $or: [
          { pdf_file_url: { $exists: false } },
          { pdf_file_url: null },
          { pdf_file_url: { $not: /r2\.cloudflarestorage\.com/ } }
        ]
      },
      {
        $set: {
          validation_status: 'archived',
          validation_notes: 'Report archived - generated before cloud storage migration. Please regenerate report.'
        }
      }
    );

    console.log(`\n📦 Archived ${result.modifiedCount} old reports`);
    console.log('✅ Users will need to regenerate these reports.\n');

  } catch (error) {
    console.error('❌ Error:', error.message);
  } finally {
    await mongoose.connection.close();
    process.exit(0);
  }
}

archiveOldReports();
