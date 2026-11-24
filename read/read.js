// Shows the open message with its LaTeX drawn. Works on the message's plain
// text (Outlook won't let an add-in change how a received message itself is
// shown, and plain text is also the safe thing to put in this page).

import { findFormulas } from "../lib/find.js";
import { renderSvg } from "../lib/render.js";
import { getText } from "../lib/mail.js";
import { reloadWhenUpdated } from "../lib/fresh.js";

const $ = (id) => document.getElementById(id);

function status(text, kind = "") {
  $("status").textContent = text;
  $("status").className = kind;
}

async function copy(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    // Some Outlook webviews block the clipboard API; fall back to a selection.
    const area = document.createElement("textarea");
    area.value = text;
    document.body.append(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }
  status(`Copied ${text.length > 40 ? text.slice(0, 40) + "..." : text}`, "ok");
}

async function show(text) {
  const box = $("message");
  box.replaceChildren();
  status("");
  const formulas = findFormulas(text);
  let at = 0;
  let bad = 0;
  for (const f of formulas) {
    let before = text.slice(at, f.start);
    // A display formula is its own block: the line breaks around it would double up.
    if (f.display) before = before.replace(/\r?\n[ \t]*$/, "");
    if (before) box.append(before);
    try {
      const svg = (await renderSvg(f.tex.replace(/ /g, " "), f.display)).cloneNode(true);
      const span = document.createElement("span");
      span.className = f.display ? "math display" : "math";
      span.title = f.source;
      span.tabIndex = 0;
      span.setAttribute("role", "button");
      span.setAttribute("aria-label", `Copy ${f.source}`);
      span.append(svg);
      span.addEventListener("click", () => copy(f.source));
      span.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          copy(f.source);
        }
      });
      box.append(span);
    } catch (e) {
      bad++;
      const span = document.createElement("span");
      span.className = "bad";
      span.title = e.message;
      span.textContent = f.source;
      box.append(span);
    }
    at = f.end;
    if (f.display) at += (/^[ \t]*\r?\n/.exec(text.slice(at)) || [""])[0].length;
  }
  if (at < text.length) box.append(text.slice(at));

  $("count").textContent = formulas.length ? `${formulas.length} formula${formulas.length === 1 ? "" : "s"}` : "";
  $("hint").hidden = !formulas.length;
  if (!formulas.length) status("No LaTeX in this message.");
  else if (bad) status(`${bad} couldn't be drawn (shown in red; hover for why).`, "error");
}

// Development: report what a received message's HTML holds (did the
// dark-mode style rule arrive? what do the formula pictures look like?).
async function reportReceived(item) {
  if (!window.__report) return;
  try {
    const html = await new Promise((resolve, reject) =>
      item.body.getAsync(Office.CoercionType.Html, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)))
    );
    const styles = html.match(/<style[\s\S]*?<\/style>/gi) || [];
    const img = (html.match(/<img[^>]*>/i) || [""])[0];
    window.__report(
      "received",
      `length=${html.length} styles=${styles.length} invert=${html.includes("invert")} ogsc=${html.includes("ogsc")} ` +
        `imgs=${(html.match(/<img/gi) || []).length}\nIMG: ${img.slice(0, 400)}\nSTYLES: ${styles.map((s) => s.slice(0, 600)).join("\n---\n")}\nTAIL: ${html.slice(-700)}`
    );
  } catch (e) {
    window.__report("received", `failed: ${e?.message || e}`);
  }
}

async function load() {
  const item = Office.context.mailbox.item;
  if (!item) return;
  reportReceived(item);
  try {
    await show(await getText(item));
  } catch (e) {
    status(`Couldn't read this message: ${e?.message || e}`, "error");
  }
}

Office.onReady(() => {
  load();
  // A pinned pane stays open as you move between messages.
  Office.context.mailbox.addHandlerAsync?.(Office.EventType.ItemChanged, load);
  reloadWhenUpdated();
});
