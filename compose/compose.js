import { scanHtml, countConverted } from "../lib/convert.js";
import { reloadWhenUpdated } from "../lib/fresh.js";
import { renderSvg } from "../lib/render.js";
import { getHtml, convertItem, revertItem, cleanUp, settings, saveSettings, appendOnSend, DARK_READER_CSS } from "../lib/mail.js";

const POLL = 700; // ms between looks at the draft
const IDLE = 1500; // ms without typing before converting

const $ = (id) => document.getElementById(id);
let busy = false;
let lastHtml = null;
let changedAt = 0;
// A draft not to auto-convert until it changes: its formulas all failed, or
// you just turned the pictures back into LaTeX to edit them.
let failedHtml = null;

function status(text, kind = "") {
  $("status").textContent = text;
  $("status").className = kind;
}

// Development: after a conversion, tell the dev server's log what Outlook kept
// of a text formula (its font list with the LaTeX marker, italic, sup/sub).
async function reportRoundTrip(item) {
  if (!window.__report) return;
  try {
    const html = await getHtml(item);
    const marked = html.match(/<[^>]*ltx[a-z0-9]{4}-[^>]*>/g) || [];
    const at = html.search(/ltx[a-z0-9]{4}-/);
    window.__report(
      "roundtrip",
      `marked=${marked.length} converted=${countConverted(html)} sup=${(html.match(/<sup/g) || []).length} ` +
        `i=${(html.match(/<i[\s>]|font-style:\s*italic/g) || []).length} imgs=${(html.match(/<img/g) || []).length} ` +
        `sample=${at >= 0 ? html.slice(Math.max(0, at - 200), at + 600) : "(no marker)"}`
    );
  } catch (e) {
    window.__report("roundtrip", `failed: ${e?.message || e}`);
  }
}

async function work(fn) {
  if (busy) return;
  busy = true;
  $("convert").disabled = $("revert").disabled = true;
  try {
    await fn();
  } catch (e) {
    status(e?.message || String(e), "error");
  } finally {
    busy = false;
    $("convert").disabled = $("revert").disabled = false;
  }
}

function report(result, auto) {
  if (result.stale) return; // you kept typing; next pause tries again
  const n = result.converted;
  const bad = result.errors.length;
  const done = n ? `Converted ${n} formula${n === 1 ? "" : "s"}.` : "";
  if (bad) status(`${done} ${bad} ${bad === 1 ? "has" : "have"} an error - see below.`.trim(), "error");
  else if (n) status(done, "ok");
  else if (!auto) status("No LaTeX found.");
}

// ---- the live list ----

let listKey = "";
async function showList(html) {
  const formulas = scanHtml(html);
  const key = formulas.map((f) => f.source).join("\u0000");
  if (key === listKey) return;
  listKey = key;

  const converted = countConverted(html);
  $("done").textContent = converted ? `${converted} converted formula${converted === 1 ? "" : "s"} in the draft.` : "";
  $("listTitle").textContent = formulas.length ? `Still LaTeX (${formulas.length})` : "Waiting for formulas";

  const list = $("list");
  list.replaceChildren();
  for (const f of formulas.slice(0, 30)) {
    const li = document.createElement("li");
    const src = document.createElement("div");
    src.className = "src";
    src.textContent = f.source;
    li.append(src);
    try {
      const svg = (await renderSvg(f.tex.replace(/ /g, " "), f.display)).cloneNode(true);
      const box = document.createElement("div");
      box.className = "math";
      box.append(svg);
      li.append(box);
    } catch (e) {
      li.className = "bad";
      const err = document.createElement("div");
      err.className = "err";
      err.textContent = e.message;
      li.append(err);
    }
    if (listKey !== key) return; // the draft changed while drawing
    list.append(li);
  }
}

// ---- the dark-mode rule for readers ----
// Added to the email when it's sent, after the signature. Outlook only lets a
// pane do this (not the send step), so it's done here, once per draft, as soon
// as the draft has a formula picture.

// OFF (2026-11-23): tested in new Outlook - the rule went out with the email
// and Outlook still showed the formula black in dark mode, so it does nothing
// there. Apple Mail might honour it, but with "light background" on in a dark
// macOS it would turn formulas white on white. Kept for reference.
const DARK_RULE_ENABLED = false;

const ruleAdded = new WeakSet();
async function ensureDarkRule(item) {
  if (!DARK_RULE_ENABLED) return;
  if (!item || ruleAdded.has(item) || !item.body?.appendOnSendAsync) return;
  ruleAdded.add(item);
  try {
    await appendOnSend(item, DARK_READER_CSS);
    window.__report?.("send", "dark-mode rule set to go out with this email (from the pane)");
  } catch (e) {
    ruleAdded.delete(item);
    window.__report?.("send", `dark-mode rule not set from the pane: ${e?.message || e}`);
  }
}

// ---- watching the draft ----

async function tick() {
  const item = Office.context.mailbox.item;
  if (!item || busy) return;
  let html;
  try {
    html = await getHtml(item);
  } catch {
    return;
  }
  const now = Date.now();
  if (html !== lastHtml) {
    lastHtml = html;
    changedAt = now;
    showList(html).catch(() => {});
    return;
  }
  if (!$("auto").checked || now - changedAt < IDLE || html === failedHtml) return;
  if (!scanHtml(html).length) return;

  await work(async () => {
    const result = await convertItem(item, { guard: true });
    report(result, true);
    if (result.converted) {
      reportRoundTrip(item);
      ensureDarkRule(item);
    }
    if (!result.converted && !result.stale) failedHtml = html;
    lastHtml = null; // re-read the new body next tick
  });
}

function loop() {
  tick().finally(() => setTimeout(loop, POLL));
}

// ---- settings ----

function bindSettings() {
  const s = settings();
  $("auto").checked = s.auto;
  $("onSend").checked = s.onSend;
  $("mode").value = s.mode === "text" ? "text" : "pictures";
  $("fontSize").value = s.fontSize;
  $("color").value = s.color;
  $("background").value = s.background;
  if ($("background").selectedIndex < 0) $("background").value = "transparent";

  const save = (key, read) => async () => {
    try {
      await saveSettings({ [key]: read() });
    } catch (e) {
      status(`Couldn't save the setting: ${e.message}`, "error");
    }
  };
  $("auto").addEventListener("change", save("auto", () => $("auto").checked));
  $("onSend").addEventListener("change", save("onSend", () => $("onSend").checked));
  $("mode").addEventListener("change", save("mode", () => $("mode").value));
  $("fontSize").addEventListener("change", save("fontSize", () => Math.min(48, Math.max(8, Number($("fontSize").value) || 16))));
  $("color").addEventListener("change", save("color", () => $("color").value));
  $("background").addEventListener("change", save("background", () => $("background").value));
}

Office.onReady(() => {
  bindSettings();
  $("convert").addEventListener("click", () =>
    work(async () => {
      const item = Office.context.mailbox.item;
      const result = await convertItem(item);
      report(result, false);
      await cleanUp(item);
      if (result.converted) {
        reportRoundTrip(item);
        ensureDarkRule(item);
      }
      lastHtml = null;
    })
  );
  $("revert").addEventListener("click", () =>
    work(async () => {
      const n = await revertItem(Office.context.mailbox.item);
      status(n ? `${n} picture${n === 1 ? "" : "s"} back to LaTeX. Edit away; they convert again once you change something.` : "No formula pictures to turn back.", n ? "ok" : "");
      lastHtml = null;
      // Don't convert them straight back: wait until the draft is edited.
      failedHtml = await getHtml(Office.context.mailbox.item);
    })
  );
  loop();
  reloadWhenUpdated(() => !busy);
  // A draft that already has formula pictures (pane opened later, or reloaded).
  getHtml(Office.context.mailbox.item)
    .then((html) => {
      if (countConverted(html)) ensureDarkRule(Office.context.mailbox.item);
      // Development: what a reply quotes of the message it answers (did the
      // dark-mode style rule arrive with it?).
      const styles = html.match(/<style[\s\S]*?<\/style>/gi) || [];
      window.__report?.(
        "draft-on-open",
        `length=${html.length} styles=${styles.length} invert=${html.includes("invert")} ogsc=${html.includes("ogsc")} ` +
          `imgs=${(html.match(/<img/gi) || []).length}\nIMG: ${((html.match(/<img[^>]*>/i) || [""])[0]).slice(0, 400)}\n` +
          `STYLES: ${styles.map((s) => s.slice(0, 700)).join("\n---\n")}`
      );
    })
    .catch(() => {});
});
