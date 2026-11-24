// Classic Outlook for Windows runs send-time handlers in a script-only runtime
// with no page to draw on, so formulas can't be redrawn there. It lets the
// message go - unless the draft holds pictures drawn in light ink for a dark
// editor (attachments named latex-xxxxxxxx-e.png), which would be near
// invisible on a recipient's white page. (New Outlook, Outlook on the web and
// Mac use commands.html instead.)

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;
  item.getAttachmentsAsync((r) => {
    const light =
      r.status === Office.AsyncResultStatus.Succeeded &&
      r.value.some((a) => /^latex-[0-9a-f]{8}-e\.png$/.test(a.name || ""));
    if (light) {
      event.completed({
        allowEvent: false,
        errorMessage:
          "LaTeX: this draft's formulas were drawn for a dark editor in new Outlook. Send it from new Outlook or Outlook on the web, or click Back to LaTeX in the LaTeX pane first.",
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

function convertNow(event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "latex",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "Open the LaTeX pane to convert formulas in this version of Outlook.",
    },
    () => event.completed()
  );
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("convertNow", convertNow);
