// Rewrites an email body: LaTeX in the text becomes inline PNG images
// (referenced as cid: attachments), and back again. Nothing here talks to
// Office - see mail.js for that - so it can be tested in a plain page.

import { findFormulas, parseSource, hash } from "./find.js";
import { renderPng, renderSvg, DEFAULTS } from "./render.js";
import { toTextMath } from "./textmath.js";

// ---- formulas written as text ----
//
// Outlook's editor keeps fonts, italic, bold, superscript and subscript, and its
// dark mode recolours text - so a formula typeset as text follows light and
// dark mode, for you and for whoever reads it. Its LaTeX rides along in the
// font list: font-family: "Cambria Math", ltxAbCd-<base64url of the LaTeX>, ...
// A font that doesn't exist is skipped by every email app, and the editor keeps
// the list on each piece of the formula, so "Back to LaTeX" can find it again.

const MATH_FONTS = `'Cambria Math', 'STIX Two Math', 'Times New Roman', serif`;
const MARK = /ltx([a-z0-9]{4})-([A-Za-z0-9_-]+)/;

function toBase64Url(s) {
  let bin = "";
  for (const b of new TextEncoder().encode(s)) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function fromBase64Url(s) {
  const bin = atob(s.replace(/-/g, "+").replace(/_/g, "/"));
  return new TextDecoder().decode(Uint8Array.from(bin, (c) => c.charCodeAt(0)));
}
const randomTag = () => Math.random().toString(36).slice(2, 6).padEnd(4, "0");

function textElement(doc, formula, html) {
  const span = doc.createElement("span");
  const marker = `ltx${randomTag()}-${toBase64Url(formula.source)}`;
  span.setAttribute("style", `font-family:${MATH_FONTS.replace("'Cambria Math',", `'Cambria Math', ${marker},`)};`);
  span.innerHTML = html;
  if (!formula.display) return span;
  const block = doc.createElement("div");
  block.setAttribute("style", "text-align:center;");
  block.append(span);
  return block;
}

/** Marker id of an element (its own style only), or null. */
const markOf = (el) => (el.getAttribute?.("style") || "").match(MARK);

/** Is every bit of text under el inside an element carrying this marker? */
function onlyFormula(el, id) {
  const walker = el.ownerDocument.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let any = false;
  while (walker.nextNode()) {
    const t = walker.currentNode;
    if (!t.data.trim()) continue;
    any = true;
    let p = t.parentElement;
    while (p && p !== el.parentElement && markOf(p)?.[0] !== id) p = p.parentElement;
    if (!p || p === el.parentElement) return false;
  }
  return any;
}

const INLINE = new Set(["SPAN", "I", "EM", "B", "STRONG", "SUP", "SUB", "FONT"]);

/** Text formulas back to LaTeX. Returns how many. */
function revertTextFormulas(doc) {
  const groups = new Map(); // marker -> outermost elements
  for (const el of doc.body.querySelectorAll("[style*='ltx']")) {
    const m = markOf(el);
    if (!m) continue;
    let up = el.parentElement;
    while (up && markOf(up)?.[0] !== m[0]) up = up.parentElement;
    if (up) continue; // inside another piece of the same formula
    if (!groups.has(m[0])) groups.set(m[0], { source: fromBase64Url(m[2]), parts: [] });
    groups.get(m[0]).parts.push(el);
  }
  for (const [id, { source, parts }] of groups) {
    // Lift each piece to the highest inline wrapper that holds only this formula.
    const lifted = parts.map((el) => {
      let top = el;
      while (top.parentElement && INLINE.has(top.parentElement.tagName) && onlyFormula(top.parentElement, id)) top = top.parentElement;
      return top;
    });
    const unique = [...new Set(lifted)];
    unique[0].replaceWith(doc.createTextNode(source));
    for (const el of unique.slice(1)) el.remove();
  }
  return groups.size;
}

const SKIP = new Set(["SCRIPT", "STYLE", "TEXTAREA", "CODE", "PRE", "TITLE"]);
const NAME_PREFIX = "latex-";

/** What Outlook's editor and autocorrect may have done to the TeX. */
function cleanTex(tex) {
  return tex
    .replace(/ /g, " ")
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/​/g, "");
}

/**
 * Attachment name of a formula's picture. Pictures drawn in light ink for a
 * dark editor (options.forEditor) end in "-e.png": they must never be sent, and
 * the send step and classic Outlook spot them by that name.
 */
export function imageName(formula, options = {}) {
  const o = { ...DEFAULTS, ...options };
  const key = [formula.display ? "D" : "I", cleanTex(formula.tex), o.fontSize, o.color, o.background].join("|");
  return `${NAME_PREFIX}${hash(key)}${o.forEditor ? "-e" : ""}.png`;
}

export const isOurAttachment = (name) => typeof name === "string" && /^latex-[0-9a-f]{8}(-e)?\.png$/.test(name);
export const isEditorInk = (name) => typeof name === "string" && /^latex-[0-9a-f]{8}-e\.png$/.test(name);

const cache = new Map();
function pngFor(formula, options) {
  const name = imageName(formula, options);
  if (!cache.has(name)) {
    const p = renderPng(cleanTex(formula.tex), formula.display, options);
    p.catch(() => cache.delete(name)); // don't keep failures
    cache.set(name, p);
  }
  return { name, png: cache.get(name) };
}

function textNodes(root) {
  const out = [];
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      for (let p = node.parentElement; p && p !== root; p = p.parentElement) {
        if (SKIP.has(p.tagName) || p.getAttribute("contenteditable") === "false") return NodeFilter.FILTER_REJECT;
      }
      return NodeFilter.FILTER_ACCEPT;
    },
  });
  while (walker.nextNode()) out.push(walker.currentNode);
  return out;
}

function parse(html) {
  return new DOMParser().parseFromString(html || "", "text/html");
}

function serialize(doc, original) {
  // Outlook hands over a whole document; give one back. A fragment stays a fragment.
  return /<html[\s>]/i.test(original) ? doc.documentElement.outerHTML : doc.body.innerHTML;
}

/** Plain-text formulas in a body, for the live list (no rendering). */
export function scanHtml(html) {
  const doc = parse(html);
  doc.body.normalize();
  return textNodes(doc.body).flatMap((n) => findFormulas(n.data));
}

/**
 * @returns {Promise<{ html: string, images: Map<string, string>, converted: number,
 *   errors: { source: string, message: string }[] }>}
 * images: attachment name -> PNG base64, for every image the new html refers to.
 */
export async function convertHtml(html, options = {}) {
  const doc = parse(html);
  doc.body.normalize();
  const images = new Map();
  const errors = [];
  let converted = 0;

  const asText = options.mode === "text";

  for (const node of textNodes(doc.body)) {
    const formulas = findFormulas(node.data);
    if (!formulas.length) continue;

    const rendered = await Promise.all(
      formulas.map(async (f) => {
        try {
          // MathJax checks every formula, so a typo is reported either way.
          await renderSvg(cleanTex(f.tex), f.display);
          const text = asText ? toTextMath(cleanTex(f.tex)) : null;
          if (text) return { f, text };
          const { name, png } = pngFor(f, options);
          return { f, name, picture: await png };
        } catch (e) {
          errors.push({ source: f.source, message: e.message || String(e) });
          return { f };
        }
      })
    );

    const frag = doc.createDocumentFragment();
    let at = 0;
    for (const { f, name, picture, text } of rendered) {
      if (!picture && !text) continue; // a formula with an error stays as text
      if (f.start > at) frag.append(node.data.slice(at, f.start));
      if (text) {
        frag.append(textElement(doc, f, text.html));
      } else {
        frag.append(imageElement(doc, f, name, picture));
        images.set(name, picture.base64);
      }
      converted++;
      at = f.end;
    }
    if (at === 0) continue;
    if (at < node.data.length) frag.append(node.data.slice(at));
    node.replaceWith(frag);
  }

  return { html: serialize(doc, html), images, converted, errors };
}

function imageElement(doc, formula, name, picture) {
  const img = doc.createElement("img");
  img.setAttribute("src", `cid:${name}`);
  img.setAttribute("alt", formula.source);
  img.setAttribute("title", formula.source);
  img.setAttribute("width", String(picture.width));
  img.setAttribute("height", String(picture.height));
  img.setAttribute(
    "style",
    `width:${picture.width}px;height:${picture.height}px;vertical-align:-${picture.baseline}px;border:0;`
  );
  if (!formula.display) return img;
  const block = doc.createElement("span");
  block.setAttribute("style", "display:block;text-align:center;margin:8px 0;");
  block.append(img);
  return block;
}

/** Our images in a body: each one's formula and the element. */
function ourImages(doc) {
  return [...doc.body.querySelectorAll("img")]
    .map((img) => ({ img, formula: parseSource(img.getAttribute("alt")) }))
    .filter(({ img, formula }) => {
      if (!formula) return false;
      const src = img.getAttribute("src") || "";
      // After Outlook re-saves the draft the cid: link may become another URL;
      // the alt text is what we can count on.
      return src.includes(NAME_PREFIX) || src.startsWith("cid:") || img.getAttribute("title") === formula.source;
    });
}

/** Turns our images back into their LaTeX source. */
export function revertHtml(html) {
  const doc = parse(html);
  const found = ourImages(doc);
  for (const { img, formula } of found) {
    const parent = img.parentElement;
    const text = doc.createTextNode(formula.source);
    const isBlock = parent && parent.tagName === "SPAN" && parent.childNodes.length === 1 && /display:\s*block/.test(parent.getAttribute("style") || "");
    (isBlock ? parent : img).replaceWith(text);
  }
  const texts = revertTextFormulas(doc);
  return { html: serialize(doc, html), reverted: found.length + texts };
}

/** How many converted formulas (pictures and text) a body holds. */
export function countConverted(html) {
  const doc = parse(html);
  const marks = new Set();
  for (const el of doc.body.querySelectorAll("[style*='ltx']")) {
    const m = markOf(el);
    if (m) marks.add(m[0]);
  }
  return ourImages(doc).length + marks.size;
}

/**
 * Attachment names the body still uses: linked by name, and the names our
 * images would have with the current style (an image made with another style
 * and relinked by Outlook can't be matched - see mail.js cleanUp).
 */
export function namesInUse(html, ...styles) {
  const linked = new Set((html || "").match(/latex-[0-9a-f]{8}(-e)?\.png/g) || []);
  // Each image: the names it could have under any of the given styles (send ink, editor ink).
  const byFormula = ourImages(parse(html)).map(({ formula }) => styles.map((s) => imageName(formula, s)));
  return { linked, byFormula };
}

/** Does the body hold any of our pictures? */
export const hasPictures = (html) => ourImages(parse(html)).length > 0;
