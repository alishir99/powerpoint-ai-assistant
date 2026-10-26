// Turns the AI's reply into a clean outline, or throws a readable error.
export function parseOutline(text) {
  // Some models wrap JSON in ``` fences even when asked not to.
  const json = text.replace(/^```(?:json)?\s*|\s*```$/g, "").trim();
  let outline;
  try {
    outline = JSON.parse(json);
  } catch {
    throw new Error("The AI did not return a valid outline. Try again.");
  }
  if (!outline || !Array.isArray(outline.slides) || outline.slides.length === 0) {
    throw new Error("The AI returned an empty outline. Try again.");
  }
  return {
    title: String(outline.title || "Presentation"),
    subtitle: String(outline.subtitle || ""),
    slides: outline.slides.map((s) => ({
      title: String(s.title || ""),
      bullets: (Array.isArray(s.bullets) ? s.bullets : []).map(String).slice(0, 6),
    })),
  };
}
