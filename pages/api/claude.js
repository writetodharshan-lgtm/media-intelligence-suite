export const config = { maxDuration: 60 };

async function callWithRetry(client, params, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      return await client.messages.create(params);
    } catch (err) {
      const isOverloaded = err?.status === 529 || err?.message?.includes("Overloaded");
      const isRateLimit = err?.status === 429;
      if ((isOverloaded || isRateLimit) && i < retries - 1) {
        // Wait before retrying: 2s, 4s, 8s
        await new Promise(r => setTimeout(r, 2000 * Math.pow(2, i)));
        continue;
      }
      throw err;
    }
  }
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { system, prompt } = req.body;
  if (!system || !prompt) {
    return res.status(400).json({ error: "system and prompt are required" });
  }

  try {
    const Anthropic = (await import("@anthropic-ai/sdk")).default;
    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

    const message = await callWithRetry(client, {
      model: "claude-sonnet-4-20250514",
      max_tokens: 2000,
      system,
      messages: [{ role: "user", content: prompt }],
    });

    const text = message.content
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join("\n");

    return res.status(200).json({ text });
  } catch (err) {
    console.error(err);
    const status = err?.status || 500;
    const msg = status === 529
      ? "Claude is temporarily overloaded. Please try again in a few seconds."
      : err.message || "API call failed";
    return res.status(status).json({ error: msg });
  }
}
