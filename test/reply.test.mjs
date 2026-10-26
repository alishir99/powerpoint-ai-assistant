import { test } from "node:test";
import assert from "node:assert/strict";
import { parseReply } from "../src/taskpane/reply.js";

test("reads a create reply", () => {
  const r = parseReply('{"action":"create","reply":"Done","title":"T","slides":[{"title":"A","bullets":["x","y"]}]}');
  assert.equal(r.action, "create");
  assert.deepEqual(r.slides[0].bullets, ["x", "y"]);
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
