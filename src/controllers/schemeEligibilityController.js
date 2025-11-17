const schemeEligibilityService = require('../services/schemeEligibilityService');

const checkEligibility = async (req, res) => {
  try {
    const formData = req.body;
    const result = await schemeEligibilityService.checkEligibility(formData);
    res.json(result);
  } catch (error) {
    console.error('Error in checkEligibility:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
};

module.exports = { checkEligibility };