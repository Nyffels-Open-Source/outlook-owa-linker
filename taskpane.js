function restId(id) {
  return Office.context.mailbox.convertToRestId(
    id,
    Office.MailboxEnums.RestVersion.v2_0
  );
}
function owaLinkFromRestId(rid) {
  return "https://outlook.office.com/mail/deeplink/read/" + encodeURIComponent(rid);
}
function parseInputToId(input) {
  if (!input) return null;
  if (/^https?:\/\//i.test(input)) {
    try {
      const url = new URL(input);
      const qId = url.searchParams.get("ItemID") || url.searchParams.get("itemid");
      if (qId) return qId;
      const parts = url.pathname.split("/").filter(Boolean);
      const idx = parts.findIndex(p => p.toLowerCase() === "read");
      if (idx >= 0 && parts[idx + 1]) return decodeURIComponent(parts[idx + 1]);
    } catch {
      return null;
    }
  }
  return input.trim();
}
function setMsg(text, cls) {
  const el = document.getElementById("msg");
  el.className = cls || "hint";
  el.textContent = text;
}
function openById(rid) {
  try {
    Office.context.mailbox.displayMessageForm(rid);
    setMsg("Opened in Outlook (if supported).", "ok");
  } catch {
    const link = owaLinkFromRestId(rid);
    window.open(link, "_blank");
    setMsg("Client not supported — opened OWA in browser.", "ok");
  }
}
async function copyCurrentOwaLink() {
  const id = Office.context.mailbox.item.itemId;
  const rid = restId(id);
  const link = owaLinkFromRestId(rid);
  if (navigator.clipboard && navigator.clipboard.writeText) {
    await navigator.clipboard.writeText(link);
  } else {
    const ta = document.createElement("textarea");
    ta.value = link; document.body.appendChild(ta); ta.select();
    document.execCommand("copy"); document.body.removeChild(ta);
  }
  setMsg("Copied current OWA link.", "ok");
}
function openCurrentHere() {
  const id = Office.context.mailbox.item.itemId;
  const rid = restId(id);
  openById(rid);
}
function onOpenClick() {
  const raw = document.getElementById("input").value;
  const id = parseInputToId(raw);
  if (!id) {
    setMsg("No valid ItemId or OWA link detected.", "err");
    return;
  }
  const rid = restId(id);
  openById(rid);
}
Office.onReady(() => {
  document.getElementById("btnOpen").addEventListener("click", onOpenClick);
  document.getElementById("btnCopyCurrent").addEventListener("click", copyCurrentOwaLink);
  document.getElementById("btnOpenCurrent").addEventListener("click", openCurrentHere);
  setMsg("Paste OWA link (any form) or REST ItemId and click Open.", "hint");
});