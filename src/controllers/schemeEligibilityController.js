const schemeEligibilityService = require('../services/schemeEligibilityService');
const logger = require('../utils/logger');

const checkEligibility = async (req, res) => {
  try {
    const formData = req.body;
    const result = await schemeEligibilityService.checkEligibility(formData);
    res.json(result);
  } catch (error) {
    logger.error('Error in checkEligibility', {
      error: error.message,
      stack: error.stack,
      operation: 'checkEligibility'
    });
    res.status(500).json({ error: 'Internal server error' });
  }
};

module.exports = { checkEligibility };