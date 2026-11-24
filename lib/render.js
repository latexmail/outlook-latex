// Turns TeX into pictures with MathJax (es5/tex-svg-full.js must be loaded first).
// renderSvg: an <svg> element for showing in our own panes.
// renderPng: a PNG (base64) plus its size and baseline, for putting in an email:
// mail clients don't run scripts and many strip SVG, but every one shows a PNG.

export const DEFAULTS = {
  fontSize: 16, // CSS px of the surrounding text (Outlook's default, Aptos 12 pt = 16 px)
  color: "#000000",
  background: "transparent", // or "#ffffff" for dark-mode readers
  scale: 3, // PNG pixels per CSS pixel, so it stays sharp on high-DPI screens
};

// How tall MathJax's "ex" is against the text size. 0.46 (a typical sans
// x-height) drew formulas too small beside Outlook's Aptos 12 pt (16 px): they
// matched only at a setting of 18, i.e. 18 x 0.46 / 16 = 0.52.
const EX_PER_EM = 0.52;

let ready = null;
export function mathjaxReady() {
  if (!ready) {
    ready = (async () => {
      for (let i = 0; i < 200 && !window.MathJax?.startup?.promise; i++) {
        await new Promise((r) => setTimeout(r, 25));
      }
      if (!window.MathJax?.startup?.promise) throw new Error("MathJax didn't load");
      await window.MathJax.startup.promise;
    })();
  }
  return ready;
}

export class TexError extends Error {}

/** @returns {Promise<SVGSVGElement>} */
export async function renderSvg(tex, display) {
  await mathjaxReady();
  const node = window.MathJax.tex2svg(tex, { display });
  const svg = node.querySelector("svg");
  const error = svg && svg.querySelector("[data-mjx-error]");
  if (!svg || error) {
    const message = error?.getAttribute("data-mjx-error") || "couldn't read this formula";
    throw new TexError(message);
  }
  return svg;
}

const exToPx = (value, exPx) => {
  const m = /^(-?[\d.]+)ex$/.exec(value || "");
  return m ? parseFloat(m[1]) * exPx : null;
};

/**
 * @returns {Promise<{ base64: string, width: number, height: number, baseline: number }>}
 * width/height in CSS px; baseline = how far the picture hangs below the text line (px).
 */
export async function renderPng(tex, display, options = {}) {
  const o = { ...DEFAULTS, ...options };
  const svg = (await renderSvg(tex, display)).cloneNode(true);
  const exPx = o.fontSize * EX_PER_EM;

  const width = exToPx(svg.getAttribute("width"), exPx);
  const height = exToPx(svg.getAttribute("height"), exPx);
  if (!width || !height) throw new TexError("couldn't measure this formula");
  const align = /vertical-align:\s*(-?[\d.]+)ex/.exec(svg.getAttribute("style") || "");
  const baseline = align ? -parseFloat(align[1]) * exPx : 0;

  svg.setAttribute("width", `${width}px`);
  svg.setAttribute("height", `${height}px`);
  svg.removeAttribute("style");
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  const markup = new XMLSerializer().serializeToString(svg).replaceAll("currentColor", o.color);

  const img = new Image();
  img.decoding = "sync";
  const loaded = new Promise((resolve, reject) => {
    img.onload = resolve;
    img.onerror = () => reject(new TexError("couldn't draw this formula"));
  });
  img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(markup);
  await loaded;

  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(width * o.scale);
  canvas.height = Math.ceil(height * o.scale);
  const ctx = canvas.getContext("2d");
  if (o.background && o.background !== "transparent") {
    ctx.fillStyle = o.background;
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  const base64 = canvas.toDataURL("image/png").split(",")[1];

  return {
    base64,
    width: Math.round(width * 10) / 10,
    height: Math.round(height * 10) / 10,
    baseline: Math.round(baseline * 10) / 10,
  };
}
