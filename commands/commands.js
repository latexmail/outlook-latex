// The ribbon's "Convert LaTeX" button and the send step.
// Runs in a hidden page (new Outlook and Outlook on the web have a full
// browser here, which drawing the PNGs needs).

import {
  convertItem,
  cleanUp,
  settings,
  saveSettings,
  prepareForSend,
  hasEditorInk,
} from "../lib/mail.js";

const log = (kind, detail) => window.__report?.(kind, detail);

function notify(item, text, error = false) {
  return new Promise((resolve) =>
    item.notificationMessages.replaceAsync(
      "latex",
      error
        ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: text }
        : {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: text,
            icon: "Icon.16x16",
            persistent: false,
          },
      resolve
    )
  );
}

const plural = (n, word) => `${n} ${word}${n === 1 ? "" : "s"}`;

function describeErrors(errors) {
  const first = errors[0];
  const more = errors.length > 1 ? ` (and ${errors.length - 1} more)` : "";
  return `${first.source.slice(0, 80)}: ${first.message}${more}`;
}

async function convertNow(event) {
  const item = Office.context.mailbox.item;
  try {
    const result = await convertItem(item);
    await cleanUp(item);
    if (result.errors.length) await notify(item, `LaTeX error in ${describeErrors(result.errors)}`.slice(0, 150), true);
    else await notify(item, result.converted ? `Converted ${plural(result.converted, "formula")}.` : "No LaTeX found in this message.");
  } catch (e) {
    await notify(item, `Couldn't convert: ${e?.message || e}`.slice(0, 150), true);
  }
  event.completed();
}

/**
 * On Send: redraw every formula picture in the sending ink (the draft may hold
 * light-ink pictures made for a dark editor), convert leftovers, and add the
 * dark-mode rule for readers whose apps honour it. Light ink must never go out:
 * if it can't be redrawn, the send stops and says why.
 */
async function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;
  const s = settings();
  log("send", "send step started");
  if (!s.sendHookSeen) {
    // From now on the editor may use light ink: this mailbox runs the send step.
    saveSettings({ sendHookSeen: true }).catch(() => {});
  }

  let editorInk = false;
  try {
    editorInk = await hasEditorInk(item);
    if (!s.onSend && !editorInk) {
      event.completed({ allowEvent: true });
      return;
    }
    const result = await prepareForSend(item);
    log("send", `converted=${result.converted} errors=${result.errors.length} editorInkBefore=${editorInk} editorInkAfter=${await hasEditorInk(item)}`);

    if (await hasEditorInk(item)) {
      event.completed({
        allowEvent: false,
        errorMessage: "LaTeX: formulas drawn for your dark editor couldn't be redrawn for sending. Click Back to LaTeX in the LaTeX pane, then send again.",
      });
      return;
    }
    if (result.errors.length) {
      event.completed({
        allowEvent: false,
        errorMessage: `A formula couldn't be drawn and would go out as plain LaTeX - ${describeErrors(result.errors)}`.slice(0, 480),
      });
      return;
    }
  } catch (e) {
    log("send", `failed: ${e?.message || e}`);
    if (editorInk) {
      event.completed({
        allowEvent: false,
        errorMessage: "LaTeX: formulas drawn for your dark editor couldn't be redrawn for sending. Click Back to LaTeX in the LaTeX pane, then send again.",
      });
      return;
    }
    // Otherwise never stand between you and sending: send as is.
  }
  event.completed({ allowEvent: true });
}

Office.actions.associate("convertNow", convertNow);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
