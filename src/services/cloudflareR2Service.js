const { S3Client, PutObjectCommand, GetObjectCommand, DeleteObjectCommand, HeadObjectCommand } = require('@aws-sdk/client-s3');
const { getSignedUrl } = require('@aws-sdk/s3-request-presigner');
const logger = require('../utils/logger');

/**
 * Cloudflare R2 Storage Service
 * Handles file uploads/downloads to/from Cloudflare R2
 */

// Initialize R2 Client (S3-compatible)
const r2Client = new S3Client({
  region: 'auto',
  endpoint: process.env.R2_ENDPOINT,
  credentials: {
    accessKeyId: process.env.R2_ACCESS_KEY_ID,
    secretAccessKey: process.env.R2_SECRET_ACCESS_KEY,
  },
});

const BUCKET_NAME = process.env.R2_BUCKET_NAME || 'dpr-excel-storage';
const R2_PUBLIC_URL = process.env.R2_PUBLIC_URL; // Optional: for public bucket access

/**
 * Upload Excel file to R2
 * @param {Object} params - Upload parameters
 * @param {Buffer} params.fileBuffer - File buffer
 * @param {string} params.userEmail - User email for folder structure
 * @param {string} params.fileName - File name
 * @returns {Promise<string>} - R2 file URL
 */
async function uploadExcel({ fileBuffer, userEmail, fileName }) {
  try {
    const key = `${userEmail}/excel/${fileName}`;
    
    const command = new PutObjectCommand({
      Bucket: BUCKET_NAME,
      Key: key,
      Body: fileBuffer,
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      Metadata: {
        uploadedAt: new Date().toISOString(),
        userEmail: userEmail,
      },
    });

    await r2Client.send(command);
    
    const fileUrl = R2_PUBLIC_URL 
      ? `${R2_PUBLIC_URL}/${key}`
      : `${process.env.R2_ENDPOINT}/${BUCKET_NAME}/${key}`;

    logger.info('Excel file uploaded to R2', {
      userEmail,
      fileName,
      key,
      fileUrl,
      size: fileBuffer.length,
    });

    return fileUrl;
  } catch (error) {
    logger.error('Error uploading Excel to R2', {
      error: error.message,
      userEmail,
      fileName,
    });
    throw new Error(`Failed to upload Excel file: ${error.message}`);
  }
}

/**
 * Upload PDF file to R2
 * @param {Object} params - Upload parameters
 * @param {Buffer} params.fileBuffer - File buffer
 * @param {string} params.userEmail - User email for folder structure
 * @param {string} params.fileName - File name
 * @returns {Promise<string>} - R2 file URL
 */
async function uploadPDF({ fileBuffer, userEmail, fileName }) {
  try {
    const key = `${userEmail}/pdf/${fileName}`;
    
    const command = new PutObjectCommand({
      Bucket: BUCKET_NAME,
      Key: key,
      Body: fileBuffer,
      ContentType: 'application/pdf',
      Metadata: {
        uploadedAt: new Date().toISOString(),
        userEmail: userEmail,
      },
    });

    await r2Client.send(command);
    
    const fileUrl = R2_PUBLIC_URL 
      ? `${R2_PUBLIC_URL}/${key}`
      : `${process.env.R2_ENDPOINT}/${BUCKET_NAME}/${key}`;

    logger.info('PDF file uploaded to R2', {
      userEmail,
      fileName,
      key,
      fileUrl,
      size: fileBuffer.length,
    });

    return fileUrl;
  } catch (error) {
    logger.error('Error uploading PDF to R2', {
      error: error.message,
      userEmail,
      fileName,
    });
    throw new Error(`Failed to upload PDF file: ${error.message}`);
  }
}

/**
 * Upload Image file to R2
 * @param {Object} params - Upload parameters
 * @param {Buffer} params.fileBuffer - File buffer
 * @param {string} params.userEmail - User email for folder structure
 * @param {string} params.fileName - File name
 * @param {string} params.contentType - Image content type
 * @returns {Promise<string>} - R2 file URL
 */
async function uploadImage({ fileBuffer, userEmail, fileName, contentType = 'image/png' }) {
  try {
    const key = `${userEmail}/images/${fileName}`;
    
    const command = new PutObjectCommand({
      Bucket: BUCKET_NAME,
      Key: key,
      Body: fileBuffer,
      ContentType: contentType,
      Metadata: {
        uploadedAt: new Date().toISOString(),
        userEmail: userEmail,
      },
    });

    await r2Client.send(command);
    
    const fileUrl = R2_PUBLIC_URL 
      ? `${R2_PUBLIC_URL}/${key}`
      : `${process.env.R2_ENDPOINT}/${BUCKET_NAME}/${key}`;

    logger.info('Image file uploaded to R2', {
      userEmail,
      fileName,
      key,
      fileUrl,
      size: fileBuffer.length,
    });

    return fileUrl;
  } catch (error) {
    logger.error('Error uploading image to R2', {
      error: error.message,
      userEmail,
      fileName,
    });
    throw new Error(`Failed to upload image: ${error.message}`);
  }
}

/**
 * Download file from R2
 * @param {string} key - File key in R2 (e.g., "user@email.com/excel/report.xlsx")
 * @returns {Promise<Buffer>} - File buffer
 */
async function downloadFile(key) {
  try {
    const command = new GetObjectCommand({
      Bucket: BUCKET_NAME,
      Key: key,
    });

    const response = await r2Client.send(command);
    
    // Convert stream to buffer
    const chunks = [];
    for await (const chunk of response.Body) {
      chunks.push(chunk);
    }
    const fileBuffer = Buffer.concat(chunks);

    logger.info('File downloaded from R2', {
      key,
      size: fileBuffer.length,
    });

    return fileBuffer;
  } catch (error) {
    logger.error('Error downloading file from R2', {
      error: error.message,
      key,
    });
    throw new Error(`Failed to download file: ${error.message}`);
  }
}

/**
 * Check if file exists in R2
 * @param {string} key - File key in R2
 * @returns {Promise<boolean>} - True if exists
 */
async function fileExists(key) {
  try {
    const command = new HeadObjectCommand({
      Bucket: BUCKET_NAME,
      Key: key,
    });

    await r2Client.send(command);
    return true;
  } catch (error) {
    if (error.name === 'NotFound') {
      return false;
    }
    throw error;
  }
}

/**
 * Delete file from R2
 * @param {string} key - File key in R2
 * @returns {Promise<void>}
 */
async function deleteFile(key) {
  try {
    const command = new DeleteObjectCommand({
      Bucket: BUCKET_NAME,
      Key: key,
    });

    await r2Client.send(command);

    logger.info('File deleted from R2', { key });
  } catch (error) {
    logger.error('Error deleting file from R2', {
      error: error.message,
      key,
    });
    throw new Error(`Failed to delete file: ${error.message}`);
  }
}

/**
 * Generate a presigned URL for downloading a file
 * @param {string} key - File key in R2
 * @param {number} expiresIn - Expiration time in seconds (default 1 hour)
 * @returns {Promise<string>} - Presigned URL
 */
async function generatePresignedUrl(key, expiresIn = 3600) {
  try {
    if (!key) return null;

    const command = new GetObjectCommand({
      Bucket: BUCKET_NAME,
      Key: key,
    });

    const signedUrl = await getSignedUrl(r2Client, command, { expiresIn });
    
    logger.info('Generated presigned URL', {
      key,
      expiresIn,
    });

    return signedUrl;
  } catch (error) {
    logger.error('Error generating presigned URL', {
      error: error.message,
      key,
    });
    return null;
  }
}

/**
 * Extract R2 key from URL
 * @param {string} url - Full R2 URL
 * @returns {string} - R2 key
 */
function extractKeyFromUrl(url) {
  if (!url) return null;
  
  try {
    // Extract key from URL like: https://endpoint.com/bucket-name/user@email.com/excel/file.xlsx
    const urlObj = new URL(url);
    const pathParts = urlObj.pathname.substring(1).split('/'); // Remove leading slash and split
    
    // If first part is bucket name, remove it
    if (pathParts[0] === BUCKET_NAME) {
      pathParts.shift();
    }
    
    const key = pathParts.join('/');
    logger.info('Extracted R2 key from URL', { url, key });
    return key;
  } catch (error) {
    logger.error('Error extracting key from URL', { url, error: error.message });
    return null;
  }
}

module.exports = {
  uploadExcel,
  uploadPDF,
  uploadImage,
  downloadFile,
  fileExists,
  deleteFile,
  extractKeyFromUrl,
  generatePresignedUrl,
};
