// Finds LaTeX formulas in plain text.
//
// Delimiters: $$...$$ and \[...\] (display), $...$ and \(...\) (inline).
// A single $ only counts as math when it can't be money: the opening $ must be
// followed by a non-space, the closing $ preceded by a non-space and not
// followed by a digit. So "$20 and $30" stays text, "$x^2$" is math.
// \$ is a literal dollar sign and never opens or closes anything.

const MAX_LENGTH = 2000;

/**
 * @typedef {{ start: number, end: number, tex: string, display: boolean, source: string }} Formula
 * start/end index the whole match including delimiters; source is that slice.
 */

function isEscaped(text, i) {
  let slashes = 0;
  for (let j = i - 1; j >= 0 && text[j] === "\\"; j--) slashes++;
  return slashes % 2 === 1;
}

const isSpace = (c) => c === undefined || /\s/.test(c);
const isDigit = (c) => c !== undefined && c >= "0" && c <= "9";

/** @returns {Formula[]} */
export function findFormulas(text) {
  const found = [];
  let i = 0;
  while (i < text.length) {
    const c = text[i];
    let match = null;

    if (c === "$" && !isEscaped(text, i)) {
      match = text[i + 1] === "$" ? matchDisplayDollar(text, i) : matchInlineDollar(text, i);
      if (!match && text[i + 1] === "$") {
        i += 2; // an unclosed $$ isn't two openers
        continue;
      }
    } else if (c === "\\" && !isEscaped(text, i)) {
      const next = text[i + 1];
      if (next === "(") match = matchBracket(text, i, "\\)", false);
      else if (next === "[") match = matchBracket(text, i, "\\]", true);
    }

    if (match) {
      found.push(match);
      i = match.end;
    } else {
      i++;
    }
  }
  return found;
}

function make(text, start, end, tex, display) {
  return { start, end, tex, display, source: text.slice(start, end) };
}

function matchDisplayDollar(text, start) {
  let j = start + 2;
  while (j < text.length - 1) {
    if (text[j] === "$" && text[j + 1] === "$" && !isEscaped(text, j)) {
      const tex = text.slice(start + 2, j).trim();
      if (!tex || tex.length > MAX_LENGTH) return null;
      return make(text, start, j + 2, tex, true);
    }
    j++;
  }
  return null;
}

function matchInlineDollar(text, start) {
  if (isSpace(text[start + 1])) return null;
  for (let j = start + 1; j < text.length && j - start <= MAX_LENGTH; j++) {
    const c = text[j];
    if (c === "\n") return null; // inline math stays on one line
    if (c !== "$" || isEscaped(text, j)) continue;
    if (text[j + 1] === "$") return null; // "$a$$" is ambiguous; leave it
    if (isSpace(text[j - 1]) || isDigit(text[j + 1])) return null;
    const tex = text.slice(start + 1, j);
    return make(text, start, j + 1, tex, false);
  }
  return null;
}

function matchBracket(text, start, close, display) {
  const end = text.indexOf(close, start + 2);
  if (end < 0) return null;
  const tex = text.slice(start + 2, end).trim();
  if (!tex || tex.length > MAX_LENGTH) return null;
  if (!display && tex.includes("\n")) return null;
  return make(text, start, end + close.length, tex, display);
}

/** Parses an image's alt text back into a formula, if it is exactly one. */
export function parseSource(alt) {
  if (typeof alt !== "string") return null;
  const trimmed = alt.trim();
  const all = findFormulas(trimmed);
  if (all.length !== 1 || all[0].start !== 0 || all[0].end !== trimmed.length) return null;
  return all[0];
}

/** Short stable hash (FNV-1a, 32 bit) as 8 hex digits. */
export function hash(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16).padStart(8, "0");
}
