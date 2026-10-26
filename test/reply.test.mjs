import { test } from "node:test";
import assert from "node:assert/strict";
import { parseReply } from "../src/taskpane/reply.js";

test("reads a create reply with all three slide types", () => {
  const r = parseReply(
    JSON.stringify({
      action: "create",
      reply: "Done",
      title: "T",
      slides: [
        { type: "points", title: "A", bullets: ["x", "y"], takeaway: "So what" },
        { type: "columns", title: "B", items: [{ heading: "h1", text: "t1" }, { heading: "h2", text: "t2" }] },
        { type: "process", title: "C", items: [{ heading: "1", text: "a" }, { heading: "2", text: "b" }] },
      ],
    })
  );
  assert.deepEqual(r.slides.map((s) => s.type), ["points", "columns", "process"]);
  assert.equal(r.slides[0].takeaway, "So what");
});

test("falls back to points when columns have fewer than two items", () => {
  const r = parseReply('{"action":"create","slides":[{"type":"columns","title":"A","items":[{"heading":"h","text":"t"}]}]}');
  assert.equal(r.slides[0].type, "points");
  assert.deepEqual(r.slides[0].bullets, ["h: t"]);
});

test("treats an unknown slide type as points", () => {
  const r = parseReply('{"action":"create","slides":[{"type":"chart","title":"A","bullets":["x"]}]}');
  assert.equal(r.slides[0].type, "points");
});

test("reads an edit reply and drops malformed edits", () => {
  const r = parseReply(
    '{"action":"edit","reply":"Shorter","edits":[{"slide":2,"shapeId":"7","text":"New"},{"slide":"x","text":1}]}'
  );
  assert.deepEqual(r.edits, [{ slide: 2, shapeId: "7", text: "New" }]);
});

test("accepts JSON wrapped in code fences", () => {
  const r = parseReply('```json\n{"action":"create","slides":[{"title":"A","bullets":["x"]}]}\n```');
  assert.equal(r.slides.length, 1);
});

test("rejects text that is not JSON", () => {
  assert.throws(() => parseReply("Sure! Here are your slides"), /usable answer/);
});

test("rejects an unknown action", () => {
  assert.throws(() => parseReply('{"action":"dance"}'), /can't use/);
});

test("rejects an edit with no valid changes", () => {
  assert.throws(() => parseReply('{"action":"edit","edits":[]}'), /any changes/);
});
