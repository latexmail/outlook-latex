// Typesets TeX as HTML text: <i> for variables, <sup>/<sub> for scripts,
// Unicode for symbols. Text, unlike a picture, is recoloured by Outlook's (and
// every other app's) dark mode, and it's the one kind of formatting Outlook's
// editor keeps. Layouts text can't do (matrices, environments, \binom...) make
// toTextMath return null, and the caller draws a picture instead.
//
// No DOM here: it runs in Node for the tests too.

class Unsupported extends Error {}
const no = (why) => {
  throw new Unsupported(why);
};

// ---- symbol tables ----

const GREEK = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ϵ", varepsilon: "ε", zeta: "ζ", eta: "η",
  theta: "θ", vartheta: "ϑ", iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π",
  varpi: "ϖ", rho: "ρ", varrho: "ϱ", sigma: "σ", varsigma: "ς", tau: "τ", upsilon: "υ", phi: "ϕ",
  varphi: "φ", chi: "χ", psi: "ψ", omega: "ω",
};
const GREEK_UPPER = {
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ", Upsilon: "Υ",
  Phi: "Φ", Psi: "Ψ", Omega: "Ω",
};

const REL = {
  le: "≤", leq: "≤", ge: "≥", geq: "≥", ne: "≠", neq: "≠", approx: "≈", equiv: "≡", sim: "∼",
  simeq: "≃", cong: "≅", propto: "∝", in: "∈", notin: "∉", ni: "∋", subset: "⊂", subseteq: "⊆",
  supset: "⊃", supseteq: "⊇", subsetneq: "⊊", to: "→", rightarrow: "→", leftarrow: "←", gets: "←",
  Rightarrow: "⇒", Leftarrow: "⇐", leftrightarrow: "↔", Leftrightarrow: "⇔", iff: "⟺",
  implies: "⟹", impliedby: "⟸", mapsto: "↦", longrightarrow: "⟶", longmapsto: "⟼",
  uparrow: "↑", downarrow: "↓", ll: "≪", gg: "≫", perp: "⊥", parallel: "∥", mid: "∣", nmid: "∤",
  models: "⊨", vdash: "⊢", prec: "≺", succ: "≻", preceq: "⪯", succeq: "⪰", asymp: "≍",
  doteq: "≐", lesssim: "≲", gtrsim: "≳", coloneqq: "≔", hookrightarrow: "↪",
  twoheadrightarrow: "↠", rightleftharpoons: "⇌",
};
const BIN = {
  pm: "±", mp: "∓", times: "×", cdot: "⋅", div: "÷", circ: "∘", ast: "∗", star: "⋆", bullet: "∙",
  oplus: "⊕", ominus: "⊖", otimes: "⊗", odot: "⊙", wedge: "∧", land: "∧", vee: "∨", lor: "∨",
  cup: "∪", cap: "∩", setminus: "∖", sqcup: "⊔", sqcap: "⊓", uplus: "⊎", dagger: "†",
};
const OP = {
  sum: "∑", prod: "∏", coprod: "∐", int: "∫", iint: "∬", iiint: "∭", oint: "∮", bigcup: "⋃",
  bigcap: "⋂", bigoplus: "⨁", bigotimes: "⨂", bigvee: "⋁", bigwedge: "⋀", bigsqcup: "⨆",
};
const ORD = {
  infty: "∞", partial: "∂", nabla: "∇", forall: "∀", exists: "∃", nexists: "∄", emptyset: "∅",
  varnothing: "∅", neg: "¬", lnot: "¬", angle: "∠", triangle: "△", hbar: "ℏ", ell: "ℓ", Re: "ℜ",
  Im: "ℑ", aleph: "ℵ", beth: "ℶ", wp: "℘", ldots: "…", dots: "…", cdots: "⋯", vdots: "⋮",
  ddots: "⋱", prime: "′", degree: "°", top: "⊤", bot: "⊥", square: "□", Box: "□", blacksquare: "■",
  checkmark: "✓", clubsuit: "♣", diamondsuit: "♢", heartsuit: "♡", spadesuit: "♠", sharp: "♯",
  flat: "♭", natural: "♮", dag: "†", ddag: "‡", S: "§", P: "¶", backslash: "\\", vert: "|",
  Vert: "‖", "|": "‖", "{": "{", "}": "}", "%": "%", $: "$", "#": "#", "&": "&", _: "_",
  imath: "ı", jmath: "ȷ", circledR: "®", copyright: "©", surd: "√", therefore: "∴",
  because: "∵", infin: "∞",
};
const OPEN = { langle: "⟨", lfloor: "⌊", lceil: "⌈", lbrace: "{", lbrack: "[", lvert: "|", lVert: "‖" };
const CLOSE = { rangle: "⟩", rfloor: "⌋", rceil: "⌉", rbrace: "}", rbrack: "]", rvert: "|", rVert: "‖" };
const FUNCS = new Set([
  "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
  "coth", "log", "ln", "lg", "exp", "lim", "limsup", "liminf", "max", "min", "sup", "inf", "det",
  "gcd", "lcm", "deg", "dim", "ker", "arg", "hom", "Pr", "tr", "rank", "sgn", "span", "id",
]);
const FUNC_TEXT = { limsup: "lim sup", liminf: "lim inf" };
const SPACES = { ",": " ", ":": " ", ";": " ", " ": " ", quad: " ", qquad: "  ", "!": "", enspace: " ", thinspace: " " };
const IGNORE = new Set([
  "displaystyle", "textstyle", "scriptstyle", "scriptscriptstyle", "limits", "nolimits",
  "big", "Big", "bigg", "Bigg", "bigl", "bigr", "Bigl", "Bigr", "biggl", "biggr", "Biggl", "Biggr",
  "middle", "nonumber", "notag", "mathstrut", "strut", "allowbreak", "relax",
]);
const ACCENTS = {
  hat: "̂", widehat: "̂", bar: "̄", tilde: "̃", widetilde: "̃",
  dot: "̇", ddot: "̈", vec: "⃗", check: "̌", acute: "́", grave: "̀",
  breve: "̆", mathring: "̊",
};
const VULGAR = {
  "1/2": "½", "1/3": "⅓", "2/3": "⅔", "1/4": "¼", "3/4": "¾", "1/5": "⅕", "2/5": "⅖", "3/5": "⅗",
  "4/5": "⅘", "1/6": "⅙", "5/6": "⅚", "1/7": "⅐", "1/8": "⅛", "3/8": "⅜", "5/8": "⅝", "7/8": "⅞",
  "1/9": "⅑", "1/10": "⅒",
};

// Letter styles. Letters with an older Unicode spot use it (wider font support).
const BB_BMP = { C: "ℂ", H: "ℍ", N: "ℕ", P: "ℙ", Q: "ℚ", R: "ℝ", Z: "ℤ" };
const CAL_BMP = { B: "ℬ", E: "ℰ", F: "ℱ", H: "ℋ", I: "ℐ", L: "ℒ", M: "ℳ", R: "ℛ", e: "ℯ", g: "ℊ", o: "ℴ" };
const FRAK_BMP = { C: "ℭ", H: "ℌ", I: "ℑ", R: "ℜ", Z: "ℨ" };
function styled(ch, base, bmp) {
  if (bmp[ch]) return bmp[ch];
  const c = ch.charCodeAt(0);
  if (c >= 65 && c <= 90) return String.fromCodePoint(base + (c - 65));
  if (c >= 97 && c <= 122) return String.fromCodePoint(base + 26 + (c - 97));
  if (base === 0x1d538 && c >= 48 && c <= 57) return String.fromCodePoint(0x1d7d8 + (c - 48)); // 𝟘..𝟡
  return no(`no letter style for ${ch}`);
}
const LETTER_STYLES = {
  mathbb: (ch) => styled(ch, 0x1d538, BB_BMP),
  mathcal: (ch) => styled(ch, 0x1d49c, CAL_BMP),
  mathscr: (ch) => styled(ch, 0x1d49c, CAL_BMP),
  mathfrak: (ch) => styled(ch, 0x1d504, FRAK_BMP),
};

// ---- tokens ----

function tokenize(src) {
  const out = [];
  for (let i = 0; i < src.length; ) {
    const c = src[i];
    if (c === "\\") {
      const m = /^[A-Za-z]+/.exec(src.slice(i + 1));
      if (m) {
        out.push({ t: "cmd", v: m[0] });
        i += 1 + m[0].length;
        if (src[i] === "*") i++; // \operatorname*, \sum* ...
      } else if (i + 1 < src.length) {
        out.push({ t: "cmd", v: src[i + 1] });
        i += 2;
      } else no("lone backslash");
    } else if (/\s/.test(c)) {
      // Spaces mean nothing in math, but \text{...} keeps them.
      if (out.at(-1)?.t !== "sp") out.push({ t: "sp", v: " " });
      i++;
    } else {
      out.push({ t: "ch", v: c });
      i++;
    }
  }
  return out;
}

const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// An atom: { k: kind, h: html, plain: text without tags (for fractions) }.
// kinds: ord, op (big operator), bin, rel, open, close, punct, fn, space.
const atom = (k, h, plain = h) => ({ k, h, plain });

class Parser {
  constructor(tokens) {
    this.tk = tokens;
    this.i = 0;
  }
  skipSpaces() {
    while (this.tk[this.i]?.t === "sp") this.i++;
  }
  peek() {
    this.skipSpaces();
    return this.tk[this.i];
  }
  next() {
    this.skipSpaces();
    return this.tk[this.i++];
  }
  /** The next token, spaces included (inside \text{...}). */
  nextRaw() {
    return this.tk[this.i++];
  }
  /** One argument: a {group} or a single token. Returns atoms. */
  arg() {
    const t = this.peek();
    if (!t) no("missing argument");
    if (t.t === "ch" && t.v === "{") {
      this.next();
      return this.list("}");
    }
    return this.list(null, 1);
  }
  /** Raw text of a {group} (for \text, \operatorname). */
  rawArg() {
    const t = this.next();
    if (!t) no("missing argument");
    if (!(t.t === "ch" && t.v === "{")) return t.t === "cmd" ? no("command in text") : t.v;
    let depth = 1;
    let s = "";
    while (this.i < this.tk.length) {
      const u = this.nextRaw();
      if (u.t === "ch" && u.v === "{") depth++;
      if (u.t === "ch" && u.v === "}" && --depth === 0) return s;
      if (u.t === "cmd") s += SPACES[u.v] ?? (/^[{}%$#&_]$/.test(u.v) ? u.v : no("command in text"));
      else s += u.v;
    }
    return no("unclosed group");
  }
  /** Optional [arg] (for \sqrt[n]). */
  optArg() {
    const t = this.peek();
    if (!(t && t.t === "ch" && t.v === "[")) return null;
    this.next();
    return this.list("]");
  }

  /** Atoms until `close` (a char) or, with count, that many atoms. */
  list(close, count = Infinity) {
    const atoms = [];
    while (atoms.length < count) {
      const t = this.peek();
      if (!t) {
        if (close) no(`missing ${close}`);
        break;
      }
      if (close && t.t === "ch" && t.v === close) {
        this.next();
        break;
      }
      this.next();
      if (t.t === "ch" && (t.v === "^" || t.v === "_")) {
        this.script(atoms, t.v);
        continue;
      }
      if (t.t === "ch" && t.v === "'") {
        // primes: x' x''
        const base = atoms.pop() || atom("ord", "");
        atoms.push({ ...base, h: base.h + "′", plain: base.plain + "′" });
        continue;
      }
      const a = t.t === "ch" ? this.char(t.v) : this.command(t.v);
      if (a) atoms.push(a);
    }
    return atoms;
  }

  script(atoms, kind) {
    const base = atoms.pop() || atom("ord", "");
    const body = join(this.arg(), true);
    if (!body.h) no("empty script");
    if (base.script === kind) no("double script");
    const tag = kind === "^" ? "sup" : "sub";
    atoms.push({ ...base, h: `${base.h}<${tag}>${body.h}</${tag}>`, plain: `${base.plain}${kind}${body.plain}`, script: kind, scripted: true });
  }

  char(c) {
    if (/[A-Za-z]/.test(c)) return atom("ord", `<i>${c}</i>`, c);
    if (/[0-9.]/.test(c)) return atom("ord", c);
    switch (c) {
      case "{":
        return { ...joined(this.list("}")), k: "ord" };
      case "+":
        return atom("bin", "+");
      case "-":
        return atom("bin", "−");
      case "*":
        return atom("bin", "∗");
      case "=":
      case "<":
      case ">":
        return atom("rel", esc(c), c);
      case ":":
        return atom("rel", ":");
      case ",":
      case ";":
        return atom("punct", c);
      case "(":
      case "[":
        return atom("open", c);
      case ")":
      case "]":
        return atom("close", c);
      case "|":
      case "/":
      case "!":
      case "?":
      case "@":
      case '"':
        return atom("ord", esc(c), c);
      case "~":
        return atom("space", " ", " ");
      case "&":
        return no("alignment");
      case "}":
        return no("unbalanced }");
      default:
        return atom("ord", esc(c), c);
    }
  }

  command(name) {
    if (GREEK[name]) return atom("ord", `<i>${GREEK[name]}</i>`, GREEK[name]);
    if (GREEK_UPPER[name]) return atom("ord", GREEK_UPPER[name]);
    if (REL[name]) return atom("rel", REL[name]);
    if (BIN[name]) return atom("bin", BIN[name]);
    if (OP[name]) return atom("op", OP[name]);
    if (ORD[name]) return atom("ord", esc(ORD[name]), ORD[name]);
    if (OPEN[name]) return atom("open", OPEN[name]);
    if (CLOSE[name]) return atom("close", CLOSE[name]);
    if (name in SPACES) return atom("space", SPACES[name], SPACES[name] ? " " : "");
    if (IGNORE.has(name)) return null;
    if (FUNCS.has(name)) return atom("fn", FUNC_TEXT[name] || name);
    if (LETTER_STYLES[name]) {
      const text = this.rawArg().replace(/\s/g, "");
      const out = [...text].map(LETTER_STYLES[name]).join("");
      return atom("ord", out);
    }
    switch (name) {
      case "frac":
      case "dfrac":
      case "tfrac":
      case "cfrac":
        return fraction(this.arg(), this.arg());
      case "sqrt": {
        const index = this.optArg();
        const body = this.arg();
        const pre = index ? `<sup>${join(index, true).h}</sup>` : "";
        return atom("ord", `${pre}√${wrap(body).h}`, `√${wrap(body).plain}`);
      }
      case "text":
      case "textrm":
      case "textnormal":
      case "mbox":
      case "textup": {
        const s = this.rawArg();
        return atom("ord", esc(s), s);
      }
      case "mathrm":
      case "operatorname":
      case "mathsf":
      case "mathtt":
      case "textsf":
      case "texttt": {
        const s = this.rawArg();
        return atom(name === "operatorname" ? "fn" : "ord", esc(s), s);
      }
      case "textit":
      case "mathit":
      case "emph": {
        const j = joined(this.arg());
        return atom("ord", `<i>${j.h.replace(/<\/?i>/g, "")}</i>`, j.plain);
      }
      case "mathbf":
      case "textbf":
      case "boldsymbol":
      case "bm":
      case "pmb": {
        const j = joined(this.arg());
        return atom("ord", `<b>${j.h}</b>`, j.plain);
      }
      case "overline": {
        const j = joined(this.arg());
        if (j.h !== esc(j.plain) && !/^<i>[^<]*<\/i>$/.test(j.h)) no("overline over formatting");
        const lined = [...j.plain].map((ch) => ch + "̅").join("");
        return atom("ord", j.h.startsWith("<i>") ? `<i>${lined}</i>` : esc(lined), lined);
      }
      case "left":
      case "right": {
        const t = this.next();
        if (!t) no("missing delimiter");
        if (t.t === "ch") return t.v === "." ? null : atom(name === "left" ? "open" : "close", esc(t.v), t.v);
        const d = OPEN[t.v] || CLOSE[t.v] || ORD[t.v] || (t.v === "{" || t.v === "}" ? t.v : null);
        return d ? atom(name === "left" ? "open" : "close", d) : no("delimiter");
      }
      case "not": {
        const t = this.next();
        const negated = { "=": "≠", in: "∉", subset: "⊄", subseteq: "⊈", equiv: "≢", sim: "≁", mid: "∤", le: "≰", ge: "≱", "<": "≮", ">": "≯" };
        const k = t && negated[t.v];
        return k ? atom("rel", k) : no("\\not");
      }
      case "pmod": {
        const j = joined(this.arg());
        return atom("ord", ` (mod ${j.h})`, ` (mod ${j.plain})`);
      }
      case "bmod":
      case "mod":
        return atom("bin", "mod");
      case "ne":
        return atom("rel", "≠");
      case "colon":
        return atom("punct", ":");
      default:
        if (ACCENTS[name]) {
          const j = joined(this.arg());
          if ([...j.plain].length !== 1) no(`\\${name} over more than one letter`);
          const marked = j.plain + ACCENTS[name];
          return atom("ord", j.h.startsWith("<i>") ? `<i>${marked}</i>` : esc(marked), marked);
        }
        return no(`\\${name}`);
    }
  }
}

/** Atoms -> one ord atom (keeps kind info for fractions/sqrt). */
function joined(atoms, script = false) {
  const j = join(atoms, script);
  return { ...atom("ord", j.h, j.plain), parts: atoms };
}

/** Parentheses round anything longer than one atom (for a/b and √(...)). */
function wrap(atoms) {
  const j = join(atoms, false);
  const single = atoms.filter((a) => a.k !== "space").length <= 1;
  return single ? j : { h: `(${j.h})`, plain: `(${j.plain})` };
}

function fraction(top, bottom) {
  const t = join(top, true);
  const b = join(bottom, true);
  const isNumber = (s) => /^\d+$/.test(s);
  if (isNumber(t.plain) && isNumber(b.plain)) {
    const v = VULGAR[`${t.plain}/${b.plain}`];
    if (v) return atom("ord", v);
    return atom("ord", `<sup>${t.h}</sup>⁄<sub>${b.h}</sub>`, `${t.plain}/${b.plain}`);
  }
  const wt = wrap(top);
  const wb = wrap(bottom);
  return atom("ord", `${wt.h}/${wb.h}`, `${wt.plain}/${wb.plain}`);
}

const NBSP = " ";
const THIN = " ";

/**
 * Joins atoms with TeX-like spacing: a space either side of a relation or
 * binary operator (none in scripts), a thin space after punctuation, round
 * big operators and function names. A binary operator with nothing on its
 * left (or after another operator) is a sign: no space ("−x").
 */
function join(atoms, script = false) {
  let h = "";
  let plain = "";
  let prev = null;
  for (let n = 0; n < atoms.length; n++) {
    let a = atoms[n];
    if (a.k === "bin" && (!prev || ["bin", "rel", "open", "punct", "op"].includes(prev.k))) a = { ...a, k: "ord" };
    let before = "";
    if (prev && !script) {
      if (a.k === "rel" || a.k === "bin") before = NBSP;
      else if (prev.k === "rel" || prev.k === "bin") before = " ";
      else if (prev.k === "punct") before = THIN;
      else if (prev.k === "op") before = THIN;
      else if (prev.k === "fn" && a.k !== "open" && a.k !== "punct" && a.k !== "close") before = THIN;
      else if (a.k === "op" || a.k === "fn") before = prev.k === "open" ? "" : THIN;
    }
    h += before + a.h;
    plain += (before ? " " : "") + a.plain;
    prev = a;
  }
  return { h: h.replace(/<\/i><i>/g, ""), plain };
}

/**
 * @returns {{ html: string } | null} null when the formula needs a picture.
 */
export function toTextMath(tex) {
  if (/\\(begin|end|binom|dbinom|tbinom|choose|over|atop|matrix|pmatrix|cases|substack|stackrel|overset|underset|overbrace|underbrace|xrightarrow|xleftarrow|color|textcolor|colorbox|boxed|fbox|cancel|bcancel|xcancel|newcommand|renewcommand|def|tag|label|ref|require|href|class|style|mathchoice|phantom|hphantom|vphantom)\b/.test(tex)) return null;
  if (/\\\\/.test(tex)) return null; // line breaks
  try {
    const p = new Parser(tokenize(tex));
    const atoms = p.list(null);
    if (p.peek()) return null;
    const { h } = join(atoms, false);
    if (!h.trim()) return null;
    return { html: h };
  } catch (e) {
    if (e instanceof Unsupported) return null;
    throw e;
  }
}
