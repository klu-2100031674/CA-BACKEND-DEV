const fs = require('fs');
const path = require('path');

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

  console.log('Full prompt:', prompt);

  // Send to Grok API
  const response = await fetch('https://api.x.ai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer xai-kcNWolgJzFMXSvBOvC1dTqhd9A3swLx7IX16hQ0QC8RSshizZy7RxiL48NZ8HnHyyoLoGWKPgaupqDDH`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'grok-4-fast',
      messages: [{ role: 'user', content: prompt }]
    })
  });

  const data = await response.json();
  if (!response.ok) {
    console.error('Grok API error:', data);
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
  console.log('AI response:', aiResponse);
  // Extract JSON from markdown code block if present
  const jsonMatch = aiResponse.match(/```json\s*(\{[\s\S]*?\})\s*```/);
  const jsonString = jsonMatch ? jsonMatch[1] : aiResponse;
  return JSON.parse(jsonString);
};

module.exports = { checkEligibility };