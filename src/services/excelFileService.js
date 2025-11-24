const ExcelFile = require('../models/ExcelFile');
const fs = require('fs').promises;
const path = require('path');
const logger = require('../utils/logger');

/**
 * Excel File Service
 * Handles storing and retrieving Excel files from MongoDB
 */
class ExcelFileService {
  /**
   * Save Excel file to MongoDB
   * @param {Object} fileData - File data object
   * @returns {Promise<Object>} - Saved file document
   */
  async saveExcelFile(fileData) {
    try {
      const {
        userId,
        templateId,
        fileName,
        filePath,
        fileBuffer,
        fileSize,
        generatedBy,
        stage = 'stage1',
        jsonData,
        allSheetsData,
        formattedWCData,
        htmlContent,
        htmlJsonData,
        pdfData,
        pdfFileName,
        meta
      } = fileData;

      let buffer = fileBuffer;

      // If filePath is provided, read the file
      if (filePath && !buffer) {
        buffer = await fs.readFile(filePath);
      }

      if (!buffer) {
        throw new Error('No file data provided');
      }

      const excelFile = new ExcelFile({
        user_id: userId,
        template_id: templateId,
        file_name: fileName,
        file_data: buffer,
        file_size: fileSize || buffer.length,
        generated_by: generatedBy,
        stage,
        json_data: jsonData,
        all_sheets_data: allSheetsData,
        formatted_wc_data: formattedWCData,
        html_content: htmlContent,
        html_json_data: htmlJsonData,
        pdf_data: pdfData,
        pdf_file_name: pdfFileName,
        meta
      });

      const savedFile = await excelFile.save();
      logger.info('Excel file saved to MongoDB', {
        operation: 'saveExcelFile',
        fileId: savedFile._id
      });

      return savedFile;
    } catch (error) {
      logger.error('Error saving Excel file', {
        operation: 'saveExcelFile',
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  /**
   * Get Excel file by ID
   * @param {string} fileId - File ID
   * @returns {Promise<Object>} - File document
   */
  async getExcelFile(fileId) {
    try {
      const file = await ExcelFile.findById(fileId).populate('user_id', 'name email').populate('generated_by', 'name email');
      if (!file) {
        throw new Error('Excel file not found');
      }
      return file;
    } catch (error) {
      logger.error('Error getting Excel file', {
        operation: 'getExcelFile',
        fileId,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  /**
   * Get Excel files by user ID
   * @param {string} userId - User ID
   * @param {Object} options - Query options
   * @returns {Promise<Array>} - Array of file documents
   */
  async getExcelFilesByUser(userId, options = {}) {
    try {
      const { page = 1, limit = 10, stage, templateId } = options;
      const query = { user_id: userId };

      if (stage) query.stage = stage;
      if (templateId) query.template_id = templateId;

      const files = await ExcelFile.find(query)
        .populate('user_id', 'name email')
        .populate('generated_by', 'name email')
        .sort({ createdAt: -1 })
        .limit(limit)
        .skip((page - 1) * limit);

      const total = await ExcelFile.countDocuments(query);

      return {
        files,
        pagination: {
          page,
          limit,
          total,
          pages: Math.ceil(total / limit)
        }
      };
    } catch (error) {
      logger.error('Error getting Excel files by user', {
        operation: 'getExcelFilesByUser',
        userId,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  /**
   * Update Excel file with AI report data
   * @param {string} fileId - File ID
   * @param {Object} aiReportData - AI report data
   * @returns {Promise<Object>} - Updated file document
   */
  async updateWithAIReport(fileId, aiReportData) {
    try {
      const { reportBuffer, reportFileName } = aiReportData;

      const updatedFile = await ExcelFile.findByIdAndUpdate(
        fileId,
        {
          ai_report_generated: true,
          ai_report_data: reportBuffer,
          ai_report_file_name: reportFileName
        },
        { new: true }
      );

      if (!updatedFile) {
        throw new Error('Excel file not found');
      }

      logger.info('AI report added to Excel file', {
        operation: 'updateWithAIReport',
        fileId
      });
      return updatedFile;
    } catch (error) {
      logger.error('Error updating Excel file with AI report', {
        operation: 'updateWithAIReport',
        fileId,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  /**
   * Increment download count
   * @param {string} fileId - File ID
   * @returns {Promise<Object>} - Updated file document
   */
  async incrementDownloadCount(fileId) {
    try {
      const updatedFile = await ExcelFile.findByIdAndUpdate(
        fileId,
        {
          $inc: { download_count: 1 },
          last_downloaded_at: new Date()
        },
        { new: true }
      );

      if (!updatedFile) {
        throw new Error('Excel file not found');
      }

      return updatedFile;
    } catch (error) {
      logger.error('Error incrementing download count', {
        operation: 'incrementDownloadCount',
        fileId,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  /**
   * Delete Excel file
   * @param {string} fileId - File ID
   * @param {string} userId - User ID (for authorization)
   * @returns {Promise<boolean>} - Success status
   */
  async deleteExcelFile(fileId, userId) {
    try {
      const file = await ExcelFile.findOne({ _id: fileId, user_id: userId });
      if (!file) {
        throw new Error('Excel file not found or access denied');
      }

      await ExcelFile.findByIdAndDelete(fileId);
      logger.info('Excel file deleted', {
        operation: 'deleteExcelFile',
        fileId,
        userId
      });

      return true;
    } catch (error) {
      logger.error('Error deleting Excel file', {
        operation: 'deleteExcelFile',
        fileId,
        userId,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  /**
   * Get file buffer for download
   * @param {string} fileId - File ID
   * @param {string} type - Type of file ('excel' or 'ai_report')
   * @returns {Promise<Object>} - File buffer and metadata
   */
  async getFileBuffer(fileId, type = 'excel') {
    try {
      const file = await ExcelFile.findById(fileId);
      if (!file) {
        throw new Error('Excel file not found');
      }

      let buffer, fileName, mimeType;

      if (type === 'ai-report' || type === 'ai_report') {
        if (!file.ai_report_data) {
          throw new Error('AI report not generated for this file');
        }
        buffer = file.ai_report_data;
        fileName = file.ai_report_file_name;
        mimeType = 'application/pdf';
      } else {
        buffer = file.file_data;
        fileName = file.file_name;
        mimeType = file.mime_type;
      }

      return {
        buffer,
        fileName,
        mimeType,
        fileSize: buffer.length
      };
    } catch (error) {
      logger.error('Error getting file buffer', { operation: 'getFileBuffer', error: error.message });
      throw error;
    }
  }
}

module.exports = new ExcelFileService();