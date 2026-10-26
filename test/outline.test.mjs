import { test } from "node:test";
import assert from "node:assert/strict";
import { parseOutline } from "../src/taskpane/outline.js";

test("reads a plain JSON outline", () => {
  const o = parseOutline('{"title":"T","subtitle":"S","slides":[{"title":"A","bullets":["x","y"]}]}');
  assert.equal(o.title, "T");
  assert.deepEqual(o.slides[0].bullets, ["x", "y"]);
});

test("accepts JSON wrapped in code fences", () => {
  const o = parseOutline('```json\n{"title":"T","slides":[{"title":"A","bullets":["x"]}]}\n```');
  assert.equal(o.slides.length, 1);
});

test("rejects text that is not JSON", () => {
  assert.throws(() => parseOutline("Sure! Here are your slides"), /valid outline/);
});

test("rejects an outline without slides", () => {
  assert.throws(() => parseOutline('{"title":"T","slides":[]}'), /empty outline/);
});
