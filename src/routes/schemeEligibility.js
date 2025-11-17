const express = require('express');
const router = express.Router();
const schemeEligibilityController = require('../controllers/schemeEligibilityController');

router.post('/check-eligibility', schemeEligibilityController.checkEligibility);

module.exports = router;