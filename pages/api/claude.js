export const config = { maxDuration: 60 };

const MODELS = [
  "claude-sonnet-4-20250514",
  "claude-haiku-4-5-20251001",
];

async function callWithFallback(client, system, prompt) {
  let lastError;

  for (const model of MODELS) {
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        return await client.messages.create({
          model,
          max_tokens: 2000,
          system,
          messages: [{ role: "user", content: prompt }],
        });
      } catch (err) {
        lastError = err;
        const isOverloaded = err?.status === 529 || err?.message?.includes("Overloaded");
        const isRateLimit = err?.status === 429;
        if (isOverloaded || isRateLimit) {
          // Wait briefly then try again or move to next model
          await new Promise(r => setTimeout(r, 1500));
          continue;
        }
        throw err; // Non-retryable error
      }
    }
  }
  throw lastError;
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

    const message = await callWithFallback(client, system, prompt);

    const text = message.content
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join("\n");

    return res.status(200).json({ text });
  } catch (err) {
    console.error(err);
    const status = err?.status || 500;
    const msg = status === 529
      ? "Claude is temporarily overloaded. Please try again in a minute."
      : err.message || "API call failed";
    return res.status(status).json({ error: msg });
  }
}
