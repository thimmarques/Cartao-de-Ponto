import { GoogleGenAI } from '@google/genai';

export const config = {
  maxDuration: 300,
};

const genAI = new GoogleGenAI(process.env.GEMINI_API_KEY || '');

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { model, contents, config: generationConfig } = req.body;

    const aiModel = genAI.getGenerativeModel({ model });
    const result = await aiModel.generateContent({
      contents,
      generationConfig,
    });

    const response = await result.response;
    const text = response.text();

    return res.status(200).json({ text });
  } catch (error: any) {
    console.error('API Error:', error);
    return res.status(500).json({ error: error.message || 'Internal Server Error' });
  }
}
