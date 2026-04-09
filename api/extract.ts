import { GoogleGenerativeAI } from '@google/generative-ai';

// Initialize with the API Key from environment variables
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY || '');

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // Validate API Key
  if (!process.env.GEMINI_API_KEY) {
    return res.status(500).json({ error: 'GEMINI_API_KEY is not configured' });
  }

  try {
    const { model: modelName, contents, config } = req.body;

    if (!contents || !Array.isArray(contents)) {
      return res.status(400).json({ error: 'Invalid or missing contents array' });
    }

    // Initialize model with requested configuration
    // Map model names if necessary (e.g., gemini-3.1-pro-preview might not be valid yet, use 1.5 pro)
    let actualModelName = modelName || 'gemini-1.5-flash';
    if (actualModelName.includes('3.1')) {
        // Fallback to 1.5 pro if 3.1 is not available or is a placeholder
        actualModelName = 'gemini-1.5-pro';
    }

    const model = genAI.getGenerativeModel({ 
      model: actualModelName,
      generationConfig: config
    });

    // Call Gemini API
    const result = await model.generateContent({
      contents: contents
    });

    const response = await result.response;
    const text = response.text();

    return res.status(200).json({ text });
  } catch (error: any) {
    console.error('API Error:', error);
    return res.status(500).json({ error: error.message || 'Internal Server Error' });
  }
}
