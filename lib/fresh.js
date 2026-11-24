// An open (pinned) pane keeps running the code it loaded, however long it
// stays open. This checks build.json every 20 s and reloads the pane when the
// add-in has been rebuilt, at a moment when it isn't busy.

const URL_ = new URL("../build.json", import.meta.url).href;

async function built() {
  try {
    const r = await fetch(URL_, { cache: "no-store" });
    return (await r.json()).built;
  } catch {
    return null;
  }
}

/** isIdle(): true when reloading now won't interrupt anything. */
export async function reloadWhenUpdated(isIdle = () => true) {
  const mine = await built();
  if (mine === null) return;
  setInterval(async () => {
    const now = await built();
    if (now !== null && now !== mine && isIdle()) location.reload();
  }, 20000);
}
