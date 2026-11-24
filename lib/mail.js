// Office.js glue: read and write the draft, its inline attachments, settings.

import { convertHtml, revertHtml, scanHtml, namesInUse, isOurAttachment, isEditorInk, hasPictures } from "./convert.js";
import { DEFAULTS } from "./render.js";

const call = (fn) =>
  new Promise((resolve, reject) =>
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)))
  );

export const getHtml = (item) => call((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
export const getText = (item) => call((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
const setHtml = (item, html) => call((cb) => item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, cb));
const listAttachments = (item) => call((cb) => item.getAttachmentsAsync(cb));
const addInline = (item, base64, name) =>
  call((cb) => item.addFileAttachmentFromBase64Async(base64, name, { isInline: true }, cb));
const removeAttachment = (item, id) => call((cb) => item.removeAttachmentAsync(id, cb));

// ---- settings (roaming: they follow the mailbox to every device) ----

export const SETTING_DEFAULTS = {
  auto: true, // convert formulas in the draft as you type (while the pane is open)
  onSend: true, // convert anything left when you press Send
  mode: "pictures", // "pictures": real TeX typesetting; "text": typeset as text where possible (follows dark mode, doesn't look like TeX)
  fontSize: DEFAULTS.fontSize,
  color: DEFAULTS.color,
  background: DEFAULTS.background,
  sendHookSeen: false, // set by the send step the first time it runs in this mailbox
};

export function settings() {
  const rs = Office.context.roamingSettings;
  const out = {};
  for (const [key, fallback] of Object.entries(SETTING_DEFAULTS)) {
    const v = rs?.get(key);
    out[key] = v === undefined || v === null ? fallback : v;
  }
  return out;
}

export function saveSettings(changes) {
  const rs = Office.context.roamingSettings;
  for (const [key, value] of Object.entries(changes)) rs.set(key, value);
  return call((cb) => rs.saveAsync(cb));
}

// A background saved by an older version (e.g. "adaptive") isn't a colour: the
// canvas would ignore it and fill the picture black. Anything unknown = see-through.
const knownBackground = (b) => (b === "transparent" || /^#[0-9a-f]{6}$/i.test(b || "") ? b : "transparent");

export const styleOf = (s) => ({
  fontSize: Number(s.fontSize) || DEFAULTS.fontSize,
  color: s.color,
  background: knownBackground(s.background),
  mode: s.mode === "text" ? "text" : "pictures",
});

// ---- ink: light while you write in a dark Outlook, black when it's sent ----
//
// A picture can't change colour with the reader's theme (Outlook's editor
// strips everything that would let it), so the add-in draws two inks at two
// times: light ink in a dark editor, so the draft reads right to you, and the
// chosen colour (black) when you press Send, so it reads right to the
// recipient. The light ink is only used once the send step has been seen to
// run in this mailbox (sendHookSeen) - otherwise a draft could go out with
// light ink on a white page.

export const EDITOR_INK = "#ececec";

function luminance(hex) {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i.exec(hex || "");
  if (!m) return null;
  const [r, g, b] = m.slice(1).map((h) => parseInt(h, 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** Is Outlook showing dark? Office's theme if it says, else the system's. */
export function outlookIsDark() {
  try {
    const l = luminance(Office.context.officeTheme?.bodyBackgroundColor);
    if (l !== null) return l < 0.4;
  } catch {}
  return !!window.matchMedia?.("(prefers-color-scheme: dark)").matches;
}

/** Style for pictures that go out with the email. */
const sendStyle = (s) => styleOf(s);
/** Style for pictures in the draft you're writing. */
function editorStyle(s) {
  const base = styleOf(s);
  if (!s.sendHookSeen || !outlookIsDark() || base.background !== "transparent") return base;
  return { ...base, color: EDITOR_INK, forEditor: true };
}

// ---- converting ----

/**
 * Converts every formula in the draft.
 * guard: give up (stale = true) if the draft changed while rendering, so
 * nothing typed meanwhile is overwritten.
 * forSend: draw in the sending ink (black); otherwise in the editor's ink.
 */
export async function convertItem(item, { guard = false, forSend = false } = {}) {
  const s = settings();
  const style = forSend ? sendStyle(s) : editorStyle(s);
  const html = await getHtml(item);
  if (!scanHtml(html).length) return { converted: 0, errors: [] };

  const result = await convertHtml(html, style);
  if (!result.converted) return result;

  const existing = new Set((await listAttachments(item)).map((a) => a.name));
  for (const [name, base64] of result.images) {
    if (!existing.has(name)) await addInline(item, base64, name);
  }

  if (guard && (await getHtml(item)) !== html) return { ...result, converted: 0, stale: true };
  await setHtml(item, result.html);
  return result;
}

export async function revertItem(item) {
  const { html, reverted } = revertHtml(await getHtml(item));
  if (reverted) await setHtml(item, html);
  await cleanUp(item);
  return reverted;
}

/** Does the draft have pictures in the editor's light ink? */
export async function hasEditorInk(item) {
  return (await listAttachments(item)).some((a) => isEditorInk(a.name));
}

/**
 * The send step: every formula picture redrawn in the sending ink (pictures
 * back to LaTeX, then converted again), leftovers converted too, unused
 * pictures removed. Returns the conversion result.
 */
export async function prepareForSend(item) {
  const html = await getHtml(item);
  let body = html;
  if (hasPictures(html)) body = revertHtml(html).html;
  if (!scanHtml(body).length) return { converted: 0, errors: [] };
  const result = await convertHtml(body, sendStyle(settings()));
  const existing = new Set((await listAttachments(item)).map((a) => a.name));
  for (const [name, base64] of result.images) {
    if (!existing.has(name)) await addInline(item, base64, name);
  }
  if (result.html !== html) await setHtml(item, result.html);
  await cleanUp(item);
  return result;
}

/**
 * Removes formula images no longer in the draft (you deleted one, or
 * reverted, or they were redrawn for sending), so recipients don't get them
 * as stray attachments.
 */
export async function cleanUp(item) {
  const html = await getHtml(item);
  const s = settings();
  const base = styleOf(s);
  const { linked, byFormula } = namesInUse(html, base, { ...base, color: EDITOR_INK, forEditor: true });
  const ours = (await listAttachments(item)).filter((a) => isOurAttachment(a.name));
  const have = new Set(ours.map((a) => a.name));
  // An image we can't match to an attachment (made with another style, its
  // link rewritten by Outlook) might be using any of them: leave them all.
  if (byFormula.some((names) => !names.some((n) => have.has(n) || linked.has(n)))) return 0;
  // Each image uses the first of its names that exists.
  const used = new Set(linked);
  for (const names of byFormula) {
    const found = names.find((n) => linked.has(n)) || names.find((n) => have.has(n));
    if (found) used.add(found);
  }
  let removed = 0;
  for (const a of ours) {
    if (!used.has(a.name)) {
      await removeAttachment(item, a.id);
      removed++;
    }
  }
  return removed;
}

// ---- dark-mode readers ----

/**
 * A style rule added after the signature when the email is sent (appendOnSend
 * doesn't go through the editor, which would strip it). Formula pictures have
 * their LaTeX as alt text, so they can be picked out: in dark mode, apps that
 * honour it flip them to light ink. [data-ogsc] is how Outlook marks text it
 * recoloured for its dark mode; prefers-color-scheme is Apple Mail and iOS.
 * An app that ignores it shows the black ink, as without it.
 */
const PICS = ['img[alt^="$"]', 'img[alt^="\\\\("]', 'img[alt^="\\\\["]'];
export const DARK_READER_CSS =
  `<style>` +
  `@media (prefers-color-scheme: dark) { ${PICS.join(", ")} { filter: invert(1) hue-rotate(180deg); } }` +
  `${PICS.map((p) => `[data-ogsc] ${p}`).join(", ")} { filter: invert(1) hue-rotate(180deg); }` +
  `</style>`;

export function appendOnSend(item, html) {
  return call((cb) => item.body.appendOnSendAsync(html, { coercionType: Office.CoercionType.Html }, cb));
}
