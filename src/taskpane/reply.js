// Turns the AI's reply into a clean "create" or "edit" action, or throws a readable error.

const TYPES = ["points", "columns", "process"];

function cleanSlide(s) {
  const bullets = (Array.isArray(s?.bullets) ? s.bullets : []).map(String).slice(0, 6);
  const items = (Array.isArray(s?.items) ? s.items : [])
    .map((item) => ({ heading: String(item?.heading || ""), text: String(item?.text || "") }))
    .filter((item) => item.heading || item.text)
    .slice(0, 4);
  let type = TYPES.includes(s?.type) ? s.type : "points";
  // Columns and processes need at least two items; otherwise show the content as points.
  if (type !== "points" && items.length < 2) type = "points";
  return {
    type,
    title: String(s?.title || ""),
    bullets: type === "points" && bullets.length === 0 ? items.map((i) => `${i.heading}: ${i.text}`) : bullets,
    items,
    takeaway: String(s?.takeaway || ""),
  };
}

export function parseReply(text) {
  // Some models wrap JSON in ``` fences even when asked not to.
  const json = String(text).replace(/^```(?:json)?\s*|\s*```$/g, "").trim();
  let reply;
  try {
    reply = JSON.parse(json);
  } catch {
    throw new Error("The AI did not return a usable answer. Try again.");
  }
  const message = String(reply?.reply || "");

  if (reply?.action === "create") {
    if (!Array.isArray(reply.slides) || reply.slides.length === 0) {
      throw new Error("The AI returned no slides. Try again.");
    }
    return {
      action: "create",
      reply: message,
      title: String(reply.title || "Presentation"),
      subtitle: String(reply.subtitle || ""),
      slides: reply.slides.slice(0, 10).map(cleanSlide),
    };
  }

  if (reply?.action === "edit") {
    const edits = (Array.isArray(reply.edits) ? reply.edits : [])
      .filter((e) => Number.isInteger(e?.slide) && typeof e?.shapeId === "string" && typeof e?.text === "string")
      .map((e) => ({ slide: e.slide, shapeId: e.shapeId, text: e.text }));
    if (edits.length === 0) throw new Error("The AI didn't suggest any changes. Try describing what to change.");
    return { action: "edit", reply: message, edits };
  }

  throw new Error("The AI gave an answer I can't use. Try rephrasing your request.");
}
