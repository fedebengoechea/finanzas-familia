export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'API key not configured' });
  }

  try {
    const { system, messages } = req.body;

    const geminiBody = {
      system_instruction: {
        parts: [{ text: system || '' }]
      },
      contents: messages.map(m => ({
        role: m.role === 'assistant' ? 'model' : 'user',
        parts: [{ text: m.content }]
      })),
      generationConfig: {
        maxOutputTokens: 1000,
        temperature: 0.7,
      }
    };

    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(geminiBody),
      }
    );

    const data = await response.json();
    console.log('Gemini response:', JSON.stringify(data).slice(0, 500));

    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    console.log('Extracted text:', text.slice(0, 300));

    return res.status(200).json({
      content: [{ type: 'text', text }]
    });

  } catch (error) {
    console.error('Error:', error.message);
    return res.status(500).json({ error: error.message });
  }
}
