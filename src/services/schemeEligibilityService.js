const fs = require('fs');
const path = require('path');
const logger = require('../utils/logger');

const knowledgeBasePath = path.join(__dirname, '../python-engine/ai_knowledge_base.json');

let knowledgeText = null;

const extractKnowledgeFromPDFs = async () => {
  if (knowledgeText) return knowledgeText;

  // Load from pre-extracted JSON (extracted from PDFs)
  const data = fs.readFileSync(knowledgeBasePath, 'utf8');
  const kb = JSON.parse(data);
  let text = '';
  for (const [key, value] of Object.entries(kb.resources)) {
    for (const page of value.pages) {
      text += page.text + '\n';
    }
  }
  knowledgeText = text;
  return text;
};

const checkEligibility = async (formData) => {
  /*
  const knowledgeText = await extractKnowledgeFromPDFs();

  // Create prompt
  const prompt = `Based on the following knowledge from Andhra Pradesh Industrial Development Policy and PMEGP guidelines:

${knowledgeText}

User details:
Name: ${formData.name}
Email: ${formData.email}
Phone: ${formData.phone}
Business Organisation: ${formData.businessOrganisation}
Business Name: ${formData.businessName}
Sector: ${formData.sector}
Place of Unit: ${formData.placeOfUnit}
Line of Activity: ${formData.lineOfActivity}
Primary Contact: ${formData.primaryContact}
Caste Category: ${formData.casteCategory}
Business Entity Type: ${formData.businessEntityType}
Investment: ${formData.investment}
State: ${formData.state}
Incentive Scheme: ${formData.incentiveScheme}

Determine eligibility for all relevant government schemes that the user may qualify for based on their details and the provided knowledge. Do not limit to only the mentioned incentive scheme; check for other eligible schemes as well.

Respond only with a valid JSON object in the exact format:
{
  "title": "Scheme Eligibility Assessment",
  "status": "Eligible" or "Not Eligible",
  "schemes": [
    {
      "name": "Scheme Name",
      "description": "Description",
      "subsidy": "Subsidy details",
      "requirements": ["req1", "req2"]
    }
  ],
  "suggestions": ["sug1", "sug2"]
}

Do not include any other text, explanations, or markdown. Only the JSON object.`;

  logger.debug('Sending prompt to Grok API', {
    promptLength: prompt.length,
    operation: 'checkEligibility'
  });

  // Send to Grok API
  const response = await fetch('https://api.x.ai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${process.env.XAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'grok-4-fast',
      messages: [{ role: 'user', content: prompt }]
    })
  });

  const data = await response.json();
  if (!response.ok) {
    logger.error('Grok API error', {
      status: response.status,
      statusText: response.statusText,
      errorData: data,
      operation: 'checkEligibility'
    });
    // Fallback to mock
    return {
      title: "Scheme Eligibility Assessment",
      status: "Eligible",
      schemes: [
        {
          name: "PMEGP",
          description: "Prime Minister's Employment Generation Programme",
          subsidy: "Up to 35% subsidy for general category, higher for special categories",
          requirements: [
            "Investment between ₹10 lakh to ₹2 crore",
            "Create employment for 2-5 persons",
            "In manufacturing or service sector",
            "Age between 18-35 years for general, 18-40 for special categories"
          ]
        },
        {
          name: "IDP 4.0",
          description: "Andhra Pradesh Industrial Development Policy 4.0",
          subsidy: "Investment subsidy up to 50% for MSMEs",
          requirements: [
            "Fixed capital investment as per policy",
            "Commence commercial production within policy period",
            "Eligible sector and location"
          ]
        }
      ],
      suggestions: [
        "Apply for PMEGP through KVIC portal",
        "Contact District Industries Centre for IDP incentives",
        "Prepare business plan and financial projections",
        "Ensure all statutory registrations are complete"
      ]
    };
  }

  const aiResponse = data.choices[0].message.content;
  logger.debug('Received AI response from Grok', {
    responseLength: aiResponse.length,
    operation: 'checkEligibility'
  });
  // Extract JSON from markdown code block if present
  const jsonMatch = aiResponse.match(/```json\s*(\{[\s\S]*?\})\s*```/);
  const jsonString = jsonMatch ? jsonMatch[1] : aiResponse;
  return JSON.parse(jsonString);
  */

  const knowledgeText = await extractKnowledgeFromPDFs();

  // Specific prompt for AP IDP 4.0 eligibility check
  const prompt = `From the industrial policy AP IDP 4.0, please tell me whether the below unit falls under ineligibility criteria and if it doesn't fall in ineligibility criteria, what are the incentives available for the unit.

Name of sole proprietorship concern: ${formData.businessName || 'Sri Krishna Aqua Farms'}

Proprietor details:
Name: ${formData.primaryContact || 'N Krishna Kumari'}
Gender: ${formData.gender || 'Female'}
Caste: ${formData.casteCategory || 'OC'}

Nature of Business: ${formData.lineOfActivity || 'Re circulated Aquaculture system (RAS)'}

Project cost:
Service Equipment - ${formData.serviceEquipment || '56.89'} Lacs
Civil works - ${formData.civilWorks || '10'} Lacs
Erection and Commissioning: ${formData.erectionCommissioning || '1.50'} Lacs

Working capital requirement - ${formData.workingCapital || '10'} Lacs

Total project cost - ${formData.totalProjectCost || '78.39'} Lacs

Based on the following knowledge from Andhra Pradesh Industrial Development Policy and PMEGP guidelines:

${knowledgeText}

Provide a detailed response in the following format:

Eligibility Assessment for [Business Name] under AP IDP 4.0

1. Does the Unit Fall Under Ineligibility Criteria?
[Yes/No, with detailed reasoning based on policy sections]

2. Incentives Available (If Eligible)
[List all applicable incentives with amounts, conditions, and calculations if eligible, or state none if ineligible]`;

  logger.debug('Sending prompt to Grok API', {
    promptLength: prompt.length,
    operation: 'checkEligibility'
  });

  // Send to Grok API
  const response = await fetch('https://api.x.ai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${process.env.GROK_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'grok-4-fast',
      messages: [{ role: 'user', content: prompt }]
    })
  });

  const data = await response.json();
  if (!response.ok) {
    logger.error('Grok API error', {
      status: response.status,
      statusText: response.statusText,
      errorData: data,
      operation: 'checkEligibility'
    });
    // Fallback response
    return {
      title: "AP IDP 4.0 Eligibility Assessment",
      status: "Under Review",
      details: "Unable to process request due to API error. Please try again later.",
      schemes: [],
      suggestions: []
    };
  }

  const aiResponse = data.choices[0].message.content;
  logger.debug('Received AI response from Grok', {
    responseLength: aiResponse.length,
    operation: 'checkEligibility'
  });

  // Return the complete AI response
  return {
    title: "AP IDP 4.0 Eligibility Assessment",
    completeResponse: aiResponse
  };
};

module.exports = { checkEligibility };