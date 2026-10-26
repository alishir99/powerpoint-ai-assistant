// The backend: holds the AI key (from .env) and asks the model what to do with the deck.
// It runs inside the dev server today; in production the same handler would run
// as an Azure Function in front of the company's Azure OpenAI.

const MAX_PROMPT = 2000;
const MAX_BODY = 100000; // characters: the request plus the text of the deck

const SYSTEM_PROMPT = `You are an assistant inside PowerPoint. You get the user's request and the
current deck as JSON: each slide has a number, whether it is selected, and its text shapes (id + text).
"This slide" or "the current slide" means the selected slide.

Reply with JSON only, in one of two shapes.

1. To create new slides, in the style of a top management consultancy:
{"action": "create", "reply": "...", "title": "...", "subtitle": "...",
 "slides": [{"type": "points" | "columns" | "process", "title": "...",
             "bullets": ["..."], "items": [{"heading": "...", "text": "..."}], "takeaway": "..."}]}
- "slides" are the content slides; the title slide is separate
- use the number of slides the user asks for; if they don't say, make 5; never more than 10
- every slide "title" is an action title: one full sentence stating the insight, max 14 words
- pick the type that fits the content, and vary types across the deck:
  "points": 3-5 "bullets", each under 14 words
  "columns": 2-4 "items" that are parallel options, drivers or pillars (heading 2-4 words, text max 25 words)
  "process": 3-4 "items" that are steps in order (heading 2-4 words, text max 20 words)
- "takeaway": the "so what" of the slide in one sentence, max 16 words
- the deck tells one storyline: situation, complication, resolution, next steps
- never invent statistics, numbers, sources or quotes; only use figures the user gives

2. To change existing slides:
{"action": "edit", "reply": "...", "edits": [{"slide": 3, "shapeId": "...", "text": "..."}]}
- only include shapes you change, using the slide numbers and shape ids from the deck
- keep one line per bullet, separated by \\n

Always:
- "reply" is one short sentence to the user saying what you did
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
      if (body.length > MAX_BODY) reject(new Error("The deck is too large to send in one request."));
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

async function handleChat(req, res) {
  // Only our own task pane may use the key, not other websites open in a browser.
  if (req.headers.origin && req.headers.origin !== `https://${req.headers.host}`) {
    return send(res, 403, { error: "Requests from other sites are not allowed." });
  }

  const { AI_BASE_URL = "https://openrouter.ai/api/v1", AI_MODEL = "openai/gpt-4o-mini", AI_API_KEY } = process.env;
  if (!AI_API_KEY) {
    return send(res, 500, { error: "The server has no AI key. Add AI_API_KEY to the .env file and restart npm start." });
  }

  let prompt, deck;
  try {
    ({ prompt, deck = [] } = await readJson(req));
  } catch (error) {
    return send(res, 400, { error: error.message });
  }
  if (typeof prompt !== "string" || !prompt.trim() || prompt.length > MAX_PROMPT) {
    return send(res, 400, { error: `Write a request of 1 to ${MAX_PROMPT} characters.` });
  }
  if (!Array.isArray(deck)) return send(res, 400, { error: "The deck must be a list of slides." });

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
          { role: "user", content: `Request: ${prompt}\n\nCurrent deck:\n${JSON.stringify(deck)}` },
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

module.exports = { handleChat };
