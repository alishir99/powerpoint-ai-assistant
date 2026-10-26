/* global document, Office, PowerPoint, fetch */

import { parseReply } from "./reply.js";

// Layout names in English and Swedish Office. If neither matches,
// fall back to the position used by the default Office theme.
const TITLE_LAYOUTS = ["Title Slide", "Rubrikbild"];
const CONTENT_LAYOUTS = ["Title and Content", "Rubrik och innehåll"];
const TEXT_SHAPES = ["Placeholder", "TextBox", "GeometricShape", "Callout"];
const MAX_TEXT = 1500; // characters per shape sent to the AI

let busy = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  const input = document.getElementById("prompt");
  document.getElementById("composer").onsubmit = (event) => {
    event.preventDefault();
    send();
  };
  input.onkeydown = (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      send();
    }
  };
  input.oninput = updateComposer;
  document.querySelectorAll(".suggestion").forEach((button) => {
    button.onclick = () => {
      input.value = button.dataset.text;
      updateComposer();
      input.focus();
    };
  });
  input.focus();
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

function updateComposer() {
  const input = document.getElementById("prompt");
  input.style.height = "auto";
  input.style.height = `${Math.min(input.scrollHeight, 160)}px`; // grow with the text
  document.getElementById("send").disabled = busy || !input.value.trim();
}

function addMessage(text, kind) {
  document.getElementById("welcome")?.remove();
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
  const text = input.value.trim();
  if (!text || busy) return;

  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    addMessage("This PowerPoint version is too old for the add-in. Update Office and try again.", "assistant error");
    return;
  }

  addMessage(text, "user");
  input.value = "";
  busy = true;
  updateComposer();
  const thinking = addMessage("Thinking…", "assistant thinking");
  try {
    const deck = await readDeck();
    const reply = await askAssistant(text, deck);
    if (reply.action === "create") {
      thinking.textContent = `Building ${reply.slides.length + 1} slides…`;
      await createSlides(reply);
      const titles = reply.slides.map((s, i) => `${i + 1}. ${s.title}`).join("\n");
      thinking.remove();
      addMessage(`${reply.reply || `Added "${reply.title}".`}\n${titles}`, "assistant");
    } else {
      thinking.textContent = "Updating your slides…";
      const changed = await applyEdits(reply.edits, deck);
      thinking.remove();
      addMessage(
        changed > 0 ? reply.reply || `Updated ${changed} text boxes.` : "I couldn't match the changes to your slides. Try again.",
        changed > 0 ? "assistant" : "assistant error"
      );
    }
  } catch (error) {
    thinking.remove();
    addMessage(error.message, "assistant error");
  } finally {
    busy = false;
    updateComposer();
    input.focus();
  }
}

// ---------- backend ----------

async function askAssistant(prompt, deck) {
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, deck }),
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `The server answered ${response.status}.`);
  return parseReply(data.content || "");
}

// ---------- PowerPoint: read ----------

// Loads the text of the given shapes. One shape without text would fail the whole
// batch, so on failure we retry one by one and skip the shapes that can't be read.
async function loadTexts(context, shapes) {
  const ranges = shapes.map((shape) => shape.textFrame.textRange);
  ranges.forEach((range) => range.load("text"));
  try {
    await context.sync();
    return ranges.map((range) => range.text);
  } catch {
    const texts = [];
    for (const shape of shapes) {
      const range = shape.textFrame.textRange;
      range.load("text");
      try {
        await context.sync();
        texts.push(range.text);
      } catch {
        texts.push(null);
      }
    }
    return texts;
  }
}

// Returns the deck as the AI sees it: [{ slide, selected, shapes: [{ id, text }] }]
async function readDeck() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const selected = context.presentation.getSelectedSlides();
    slides.load("items/id");
    selected.load("items/id");
    await context.sync();
    const selectedIds = new Set(selected.items.map((s) => s.id));

    slides.items.forEach((slide) => slide.shapes.load("items/id,items/type"));
    await context.sync();

    const deck = [];
    for (const [i, slide] of slides.items.entries()) {
      const shapes = slide.shapes.items.filter((shape) => TEXT_SHAPES.includes(shape.type));
      const texts = await loadTexts(context, shapes);
      deck.push({
        slide: i + 1,
        selected: selectedIds.has(slide.id),
        shapes: shapes
          .map((shape, j) => ({ id: shape.id, text: texts[j] }))
          .filter((s) => s.text !== null)
          .map((s) => ({ id: s.id, text: s.text.slice(0, MAX_TEXT) })),
      });
    }
    return deck;
  });
}

// ---------- PowerPoint: edit ----------

async function applyEdits(edits, deck) {
  // Only touch shapes we actually sent: the AI can't edit anything it wasn't shown.
  const known = new Set(deck.flatMap((s) => s.shapes.map((shape) => `${s.slide}:${shape.id}`)));
  const valid = edits.filter((e) => known.has(`${e.slide}:${e.shapeId}`));
  if (valid.length === 0) return 0;

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    valid.forEach((e) => {
      slides.items[e.slide - 1].shapes.getItem(e.shapeId).textFrame.textRange.text = e.text;
    });
    await context.sync();
    context.presentation.setSelectedSlides([slides.items[valid[0].slide - 1].id]);
    await context.sync();
  });
  return valid.length;
}

// ---------- PowerPoint: create ----------

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
  const texts = await loadTexts(context, shapes.items);
  return texts.every((text) => text !== null && text.trim() === "");
}

function fill(slide, texts) {
  // New slides come with the layout's placeholders: [0] title, [1] body or subtitle.
  texts.forEach((text, i) => {
    const shape = slide.shapes.items[i];
    if (shape) shape.textFrame.textRange.text = text;
  });
}

async function createSlides(outline) {
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
    added[0].load("id");
    await context.sync();
    context.presentation.setSelectedSlides([added[0].id]);
    await context.sync();
  });
}
