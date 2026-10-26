/* global document, Office, PowerPoint, localStorage, fetch */

import { parseOutline } from "./outline.js";

const DEFAULTS = {
  baseUrl: "https://openrouter.ai/api/v1",
  model: "openai/gpt-4o-mini",
};

// Layout names in English and Swedish Office. If neither matches,
// fall back to the position used by the default Office theme.
const TITLE_LAYOUTS = ["Title Slide", "Rubrikbild"];
const CONTENT_LAYOUTS = ["Title and Content", "Rubrik och innehåll"];

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("generate").onclick = generate;
  document.getElementById("save-settings").onclick = saveSettings;
  loadSettings();
});

// ---------- settings (kept in this computer's PowerPoint only) ----------

function readStore(key) {
  try {
    return localStorage.getItem(key) || "";
  } catch {
    return "";
  }
}

function writeStore(key, value) {
  try {
    localStorage.setItem(key, value);
  } catch {
    // storage blocked: settings only last for this session
  }
}

function loadSettings() {
  document.getElementById("base-url").value = readStore("baseUrl") || DEFAULTS.baseUrl;
  document.getElementById("model").value = readStore("model") || DEFAULTS.model;
  document.getElementById("api-key").value = readStore("apiKey");
  if (!readStore("apiKey")) document.getElementById("settings").open = true;
}

function saveSettings() {
  writeStore("baseUrl", document.getElementById("base-url").value.trim());
  writeStore("model", document.getElementById("model").value.trim());
  writeStore("apiKey", document.getElementById("api-key").value.trim());
  setStatus("Settings saved.");
}

function getSettings() {
  return {
    baseUrl: (document.getElementById("base-url").value.trim() || DEFAULTS.baseUrl).replace(/\/+$/, ""),
    model: document.getElementById("model").value.trim() || DEFAULTS.model,
    apiKey: document.getElementById("api-key").value.trim(),
  };
}

function setStatus(message, isError = false) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = isError ? "error" : "";
}

// ---------- main flow ----------

async function generate() {
  const prompt = document.getElementById("prompt").value.trim();
  const count = Number(document.getElementById("count").value);
  const settings = getSettings();

  if (!prompt) return setStatus("Write what the presentation should be about.", true);
  if (!settings.apiKey) {
    document.getElementById("settings").open = true;
    return setStatus("Add your API key under Settings first.", true);
  }
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
    return setStatus("This PowerPoint version is too old for the add-in. Update Office and try again.", true);
  }

  const button = document.getElementById("generate");
  button.disabled = true;
  try {
    setStatus("Asking the AI for an outline…");
    const outline = await askForOutline(prompt, count, settings);
    setStatus(`Building ${outline.slides.length + 1} slides…`);
    await buildSlides(outline);
    setStatus(`Done: added ${outline.slides.length + 1} slides.`);
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

// ---------- AI ----------

function systemPrompt(count) {
  return `You write PowerPoint presentations. Reply with JSON only, in exactly this shape:
{"title": "...", "subtitle": "...", "slides": [{"title": "...", "bullets": ["...", "..."]}]}
Rules:
- exactly ${count} items in "slides" (the title slide is separate)
- 3 to 5 bullets per slide, each under 12 words, no full stop at the end
- concrete and specific, no filler
- write in the same language as the user's request`;
}

async function askForOutline(prompt, count, { baseUrl, model, apiKey }) {
  // Azure OpenAI uses an "api-key" header; OpenAI-compatible APIs use a bearer token.
  const auth = baseUrl.includes("azure.com") ? { "api-key": apiKey } : { Authorization: `Bearer ${apiKey}` };

  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...auth },
    body: JSON.stringify({
      model,
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: systemPrompt(count) },
        { role: "user", content: prompt },
      ],
    }),
  });
  if (!response.ok) {
    const detail = (await response.text()).slice(0, 200);
    throw new Error(`The AI request failed (${response.status}). ${detail}`);
  }
  const data = await response.json();
  return parseOutline(data.choices?.[0]?.message?.content ?? "");
}

// ---------- PowerPoint ----------

function findLayout(layouts, names, fallbackIndex) {
  return (
    layouts.find((l) => names.includes(l.name)) || layouts[Math.min(fallbackIndex, layouts.length - 1)]
  );
}

async function isBlankDeck(context) {
  // A new presentation starts with one empty title slide. If that's all there is,
  // we remove it after building, so the deck starts with our title slide.
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  if (slides.items.length !== 1) return false;

  const shapes = slides.items[0].shapes;
  shapes.load("items");
  await context.sync();
  try {
    const ranges = shapes.items.map((shape) => {
      const range = shape.textFrame.textRange;
      range.load("text");
      return range;
    });
    await context.sync();
    return ranges.every((range) => range.text.trim() === "");
  } catch {
    return false; // a shape without text (e.g. a picture): leave the slide alone
  }
}

function fill(slide, texts) {
  // New slides come with the layout's placeholders: [0] title, [1] body or subtitle.
  texts.forEach((text, i) => {
    const shape = slide.shapes.items[i];
    if (shape) shape.textFrame.textRange.text = text;
  });
}

async function buildSlides(outline) {
  await PowerPoint.run(async (context) => {
    const removeFirst = await isBlankDeck(context);

    const layouts = context.presentation.slideMasters.getItemAt(0).layouts;
    layouts.load("items/id,items/name");
    await context.sync();
    const titleLayout = findLayout(layouts.items, TITLE_LAYOUTS, 0);
    const contentLayout = findLayout(layouts.items, CONTENT_LAYOUTS, 1);

    // 1. Add the slides (they are appended at the end of the deck).
    context.presentation.slides.add({ layoutId: titleLayout.id });
    outline.slides.forEach(() => context.presentation.slides.add({ layoutId: contentLayout.id }));
    await context.sync();

    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    const added = slides.items.slice(-(outline.slides.length + 1));
    added.forEach((slide) => slide.shapes.load("items"));
    await context.sync();

    // 2. Fill in the text.
    fill(added[0], [outline.title, outline.subtitle]);
    outline.slides.forEach((s, i) => fill(added[i + 1], [s.title, s.bullets.join("\n")]));
    await context.sync();

    // 3. Tidy up: drop the empty starter slide and jump to the title slide.
    if (removeFirst) {
      slides.items[0].delete();
      await context.sync();
    }
    if (Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
      added[0].load("id");
      await context.sync();
      context.presentation.setSelectedSlides([added[0].id]);
      await context.sync();
    }
  });
}
