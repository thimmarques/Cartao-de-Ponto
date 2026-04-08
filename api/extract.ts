import { GoogleGenerativeAI } from '@google/genai';

export const config = {
  maxDuration: 300,
};

// Initialize with the correct options object
const genAI = new GoogleGenerativeAI({ apiKey: process.env.GEMINI_API_KEY || '' });

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // Validate API Key
  if (!process.env.GEMINI_API_KEY) {
    return res.status(500).json({ error: 'GEMINI_API_KEY is not configured' });
  }

  try {
    const { files, extractionType } = req.body;

    // Validate inputs
    if (!files || !Array.isArray(files) || files.length === 0) {
      return res.status(400).json({ error: 'Invalid or missing files array' });
    }

    // Initialize model with requested configuration
    const model = genAI.getGenerativeModel({ 
      model: 'gemini-1.5-flash', // Using a stable model
      generationConfig: {
        temperature: 0.4,
        topK: 40,
        topP: 0.95,
        maxOutputTokens: 8192,
      }
    });

    // Process files (simplified example, assuming files contain base64 data and mimeType)
    const results = await Promise.all(files.map(async (file: any) => {
      const result = await model.generateContent([
        `Extract data for ${extractionType}`,
        {
          inlineData: {
            data: file.data,
            mimeType: file.mimeType
          }
        }
      ]);
      return result.response.text();
    }));

    return res.status(200).json({ results });
  } catch (error: any) {
    console.error('API Error:', error);
    return res.status(500).json({ error: error.message || 'Internal Server Error' });
  }
}
