export default async function handler(req, res) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'No API key' });

  const { system, messages } = req.body;
  const lastMessage = messages?.[messages.length - 1]?.content || '';
  const prompt = system ? `${system}\n\nMENSAJE: ${lastMessage}` : lastMessage;

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 1000, temperature: 0.7 }
      }),
    });
    const data = await response.json();
    // Return full response for debugging
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    return res.status(200).json({ content: [{ type: 'text', text }], _debug: data });
  } catch(e) {
    return res.status(500).json({ error: e.message });
  }
}
