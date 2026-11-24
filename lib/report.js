// Development only: when served from localhost, sends this page's errors and
// a "loaded" note to the dev server's /log, which prints them. Lets us see what
// happens inside Outlook, where there's no console to look at.
(() => {
  if (location.hostname !== "localhost") return;
  const page = location.pathname;
  const send = (kind, detail) => {
    try {
      fetch("/log", {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({ page, kind, detail: String(detail).slice(0, 2000) }),
        keepalive: true,
      }).catch(() => {});
    } catch {}
  };
  window.__report = send;
  window.addEventListener("error", (e) =>
    send("error", `${e.message} @ ${e.filename}:${e.lineno}:${e.colno}${e.error?.stack ? "\n" + e.error.stack : ""}`)
  );
  window.addEventListener("unhandledrejection", (e) => send("rejection", e.reason?.stack || e.reason?.message || e.reason));
  for (const level of ["error", "warn"]) {
    const original = console[level];
    console[level] = (...args) => {
      send(level, args.map((a) => (a?.stack || (typeof a === "object" ? JSON.stringify(a) : a))).join(" "));
      original.apply(console, args);
    };
  }
  send("loaded", navigator.userAgent);
  const ready = () => {
    if (!window.Office?.onReady) return send("error", "office.js did not load");
    Office.onReady((info) => {
      const item = Office.context?.mailbox?.item;
      send(
        "office-ready",
        `host=${info?.host} platform=${info?.platform} mailbox1.12=${Office.context?.requirements?.isSetSupported?.("Mailbox", "1.12")} item=${item ? item.itemType : "none"} compose=${!!item?.body?.setAsync} mathjax=${!!window.MathJax?.tex2svg} ` +
          `theme=${(() => { try { return JSON.stringify(Office.context.officeTheme); } catch { return "n/a"; } })()} prefersDark=${window.matchMedia?.("(prefers-color-scheme: dark)").matches} appendOnSend=${!!item?.body?.appendOnSendAsync} sendHookSeen=${Office.context.roamingSettings?.get?.("sendHookSeen")}`
      );
    });
  };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", ready);
  else ready();
})();
