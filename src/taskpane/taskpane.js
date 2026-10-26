/* global document, Office, PowerPoint, fetch */

import { parseOutline } from "./outline.js";

// Layout names in English and Swedish Office. If neither matches,
// fall back to the position used by the default Office theme.
const TITLE_LAYOUTS = ["Title Slide", "Rubrikbild"];
const CONTENT_LAYOUTS = ["Title and Content", "Rubrik och innehåll"];

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  const prompt = document.getElementById("prompt");
  document.getElementById("composer").onsubmit = (event) => {
    event.preventDefault();
    send();
  };
  prompt.onkeydown = (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      send();
    }
  };
  prompt.focus();
  openWithThisDeck();
});

// Remember the assistant in this deck, so the chat opens by itself next time it's opened.
function openWithThisDeck() {
  try {
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();
  } catch {
    // not supported here: the AI Slides button still works
  }
}

// ---------- chat ----------

function addMessage(text, kind) {
  const messages = document.getElementById("messages");
  const el = document.createElement("div");
  el.className = `msg ${kind}`;
  el.textContent = text;
  messages.appendChild(el);
  messages.scrollTop = messages.scrollHeight;
  return el;
}

async function send() {
  const input = document.getElementById("prompt");
  const button = document.getElementById("send");
  const text = input.value.trim();
  if (!text || button.disabled) return;

  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
    addMessage("This PowerPoint version is too old for the add-in. Update Office and try again.", "assistant error");
    return;
  }

  addMessage(text, "user");
  input.value = "";
  button.disabled = true;
  const thinking = addMessage("Writing your slides…", "assistant thinking");
  try {
    const outline = await askForOutline(text);
    thinking.textContent = `Building ${outline.slides.length + 1} slides…`;
    await buildSlides(outline);
    thinking.remove();
    const titles = outline.slides.map((s, i) => `${i + 1}. ${s.title}`).join("\n");
    addMessage(`Added "${outline.title}" with ${outline.slides.length} slides:\n${titles}`, "assistant");
  } catch (error) {
    thinking.remove();
    addMessage(error.message, "assistant error");
  } finally {
    button.disabled = false;
    input.focus();
  }
}

// ---------- backend ----------

async function askForOutline(prompt) {
  const response = await fetch("/api/outline", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt }),
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `The server answered ${response.status}.`);
  return parseOutline(data.content || "");
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
