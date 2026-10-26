// The backend: holds the AI key (from .env) and asks the model for a slide outline.
// It runs inside the dev server today; in production the same handler would run
// as an Azure Function in front of the company's Azure OpenAI.

const MAX_PROMPT = 2000;

const SYSTEM_PROMPT = `You write PowerPoint presentations. Reply with JSON only, in exactly this shape:
{"title": "...", "subtitle": "...", "slides": [{"title": "...", "bullets": ["...", "..."]}]}
Rules:
- "slides" are the content slides; the title slide is separate
- use the number of slides the user asks for; if they don't say, make 5; never more than 10
- 3 to 5 bullets per slide, each under 12 words, no full stop at the end
- concrete and specific, no filler
- write in the same language as the user's request`;

function send(res, status, data) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json");
  res.end(JSON.stringify(data));
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += chunk;
      if (body.length > 10000) reject(new Error("The request is too large."));
    });
    req.on("end", () => {
      try {
        resolve(JSON.parse(body || "{}"));
      } catch {
        reject(new Error("The request is not valid JSON."));
      }
    });
    req.on("error", reject);
  });
}

async function handleOutline(req, res) {
  // Only our own task pane may use the key, not other websites open in a browser.
  if (req.headers.origin && req.headers.origin !== `https://${req.headers.host}`) {
    return send(res, 403, { error: "Requests from other sites are not allowed." });
  }

  const { AI_BASE_URL = "https://openrouter.ai/api/v1", AI_MODEL = "openai/gpt-4o-mini", AI_API_KEY } = process.env;
  if (!AI_API_KEY) {
    return send(res, 500, { error: "The server has no AI key. Add AI_API_KEY to the .env file and restart npm start." });
  }

  let prompt;
  try {
    ({ prompt } = await readJson(req));
  } catch (error) {
    return send(res, 400, { error: error.message });
  }
  if (typeof prompt !== "string" || !prompt.trim() || prompt.length > MAX_PROMPT) {
    return send(res, 400, { error: `Write a request of 1 to ${MAX_PROMPT} characters.` });
  }

  const baseUrl = AI_BASE_URL.replace(/\/+$/, "");
  // Azure OpenAI uses an "api-key" header; OpenAI-compatible APIs use a bearer token.
  const auth = baseUrl.includes("azure.com") ? { "api-key": AI_API_KEY } : { Authorization: `Bearer ${AI_API_KEY}` };

  try {
    const response = await fetch(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...auth },
      body: JSON.stringify({
        model: AI_MODEL,
        response_format: { type: "json_object" },
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: prompt },
        ],
      }),
    });
    if (!response.ok) {
      const detail = (await response.text()).slice(0, 200);
      return send(res, 502, { error: `The AI service answered ${response.status}. ${detail}` });
    }
    const data = await response.json();
    return send(res, 200, { content: data.choices?.[0]?.message?.content ?? "" });
  } catch (error) {
    return send(res, 502, { error: `Could not reach the AI service: ${error.message}` });
  }
}

module.exports = { handleOutline };
