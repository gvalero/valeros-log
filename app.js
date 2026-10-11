// Register an app (see README.md) and paste its Application (client) ID here.
const CLIENT_ID = "b72347df-34d5-463d-8718-40af9135a64b";

const SCOPES = ["User.Read", "Files.ReadWrite.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const FILE_PATH = "ValerosLog/valeros-log.json";
const REF_KEY = "valeros-log-file-ref";
const SCHEMA_VERSION = 2;
const MAX_VISIBLE = 10;

const $ = (selector) => document.querySelector(selector);
const gate = $("#gate");
const setup = $("#setup");
const appEl = $("#app");
const form = $("#incident-form");
const formMessage = $("#form-message");
const saveButton = $("#save-entry");
const entryList = $("#entry-list");
const emptyState = $("#empty-state");
const storageStatus = $("#storage-status");

let msal = null;
let account = null;
let fileRef = null;
let entries = [];

function say(element, message, isError = false) {
  element.textContent = message;
  element.classList.toggle("error", isError);
}

function show(which) {
  gate.hidden = which !== "gate";
  setup.hidden = which !== "setup";
  appEl.hidden = which !== "app";
  $("#sign-out").hidden = !account;
  $("#who").hidden = !account;
}

async function initAuth() {
  if (!CLIENT_ID) {
    $("#gate-text").textContent =
      "This site isn't connected to Microsoft yet. Add the app's client ID in app.js (see README.md).";
    $("#sign-in").disabled = true;
    return show("gate");
  }
  if (!window.msal) {
    $("#gate-text").textContent = "Couldn't load the Microsoft sign-in library. Check your connection and reload.";
    $("#sign-in").disabled = true;
    return show("gate");
  }

  msal = new window.msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/consumers",
      redirectUri: location.origin + location.pathname,
    },
    cache: { cacheLocation: "localStorage" },
  });
  await msal.initialize();
  const result = await msal.handleRedirectPromise();
  account = result?.account || msal.getAllAccounts()[0] || null;
  if (!account) return show("gate");
  $("#who").textContent = account.username;
  await start();
}

async function getToken() {
  try {
    return (await msal.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
  } catch (error) {
    if (error instanceof window.msal.InteractionRequiredAuthError) {
      await msal.acquireTokenRedirect({ scopes: SCOPES, account });
    }
    throw error;
  }
}

async function graph(path, options = {}) {
  const response = await fetch(path.startsWith("http") ? path : GRAPH + path, {
    ...options,
    headers: { Authorization: `Bearer ${await getToken()}`, ...options.headers },
  });
  if (!response.ok) {
    const error = new Error(`Microsoft returned ${response.status}`);
    error.status = response.status;
    throw error;
  }
  return response;
}

const itemPath = () => `/drives/${fileRef.driveId}/items/${fileRef.itemId}`;

function parseLog(text) {
  if (!text.trim()) return [];
  const data = JSON.parse(text);
  if (!Array.isArray(data.entries)) throw new Error("The log file isn't in the expected format.");
  return data.entries;
}

async function loadLog() {
  const meta = await (await graph(itemPath())).json();
  const text = await (await graph(`${itemPath()}/content`)).text();
  return { entries: parseLog(text), eTag: meta.eTag };
}

function serialize(list) {
  return JSON.stringify({ schemaVersion: SCHEMA_VERSION, updatedAt: new Date().toISOString(), entries: list }, null, 2);
}

function rememberRef(item) {
  fileRef = { driveId: item.parentReference.driveId, itemId: item.id };
  localStorage.setItem(REF_KEY, JSON.stringify(fileRef));
}

async function start() {
  try {
    fileRef = JSON.parse(localStorage.getItem(REF_KEY));
  } catch {
    fileRef = null;
  }
  if (!fileRef) {
    try {
      // Reuse the log if it already exists in this account's OneDrive.
      rememberRef(await (await graph(`/me/drive/root:/${FILE_PATH}`)).json());
    } catch (error) {
      if (error.status !== 404) console.error(error);
    }
  }
  if (!fileRef) return show("setup");
  show("app");
  await refresh();
}

async function createLog() {
  const button = $("#create-log");
  button.disabled = true;
  try {
    const response = await graph(`/me/drive/root:/${FILE_PATH}:/content?@microsoft.graph.conflictBehavior=fail`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: serialize([]),
    });
    rememberRef(await response.json());
    show("app");
    await refresh();
  } catch (error) {
    say($("#setup-message"), "Couldn't create the log. Check your connection and try again.", true);
  } finally {
    button.disabled = false;
  }
}

async function connectByLink(event) {
  event.preventDefault();
  const link = $("#share-link").value.trim();
  if (!link) return;
  const encoded = "u!" + btoa(link).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  try {
    rememberRef(await (await graph(`/shares/${encoded}/driveItem`)).json());
    show("app");
    await refresh();
  } catch {
    say($("#setup-message"), "Couldn't open that link. Check it's the log's share link and allows editing.", true);
  }
}

function timeText(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime())
    ? ""
    : new Intl.DateTimeFormat(undefined, { weekday: "short", day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }).format(date);
}

function render() {
  entryList.replaceChildren();
  emptyState.hidden = entries.length > 0;
  const recent = [...entries].sort((a, b) => new Date(b.occurredAt) - new Date(a.occurredAt)).slice(0, MAX_VISIBLE);
  for (const entry of recent) {
    const item = document.createElement("li");
    item.className = "entry";
    const time = document.createElement("time");
    time.dateTime = entry.occurredAt;
    time.textContent = [timeText(entry.occurredAt), entry.who, entry.durationLabel].filter(Boolean).join(" · ");
    const summary = document.createElement("p");
    summary.textContent = entry.whatHappened;
    item.append(time, summary);
    if (entry.triggers?.length) {
      const trigger = document.createElement("p");
      trigger.className = "muted";
      trigger.textContent = `Set off by: ${entry.triggers.join(", ")}`;
      item.append(trigger);
    }
    if (entry.helped?.length) {
      const helped = document.createElement("p");
      helped.className = "muted";
      helped.textContent = `Helped: ${entry.helped.join(", ")}`;
      item.append(helped);
    }
    if (entry.note) {
      const note = document.createElement("p");
      note.textContent = entry.note;
      item.append(note);
    }
    entryList.append(item);
  }
}

async function refresh() {
  say(storageStatus, "Loading…");
  try {
    entries = (await loadLog()).entries;
    render();
    say(storageStatus, `${entries.length} saved in your OneDrive log.`);
  } catch (error) {
    if (error.status === 404) {
      localStorage.removeItem(REF_KEY);
      fileRef = null;
      return show("setup");
    }
    say(storageStatus, "Couldn't load the log. Pull to refresh or try again.", true);
  }
}

const DURATION_LABELS = { 5: "under 10 min", 20: "10–30 min", 45: "30–60 min", 75: "60+ min" };

function checked(group) {
  const other = form.querySelector(`[data-other="${group}"]`)?.value.trim() || "Other";
  return [...form.querySelectorAll(`[data-group="${group}"] input:checked`)].map((input) =>
    input.value === "__other" ? other : input.value,
  );
}

function syncOtherFields() {
  for (const field of form.querySelectorAll("[data-other]")) {
    const on = form.querySelector(`[data-group="${field.dataset.other}"] input[value="__other"]`).checked;
    field.hidden = !on;
    if (!on) field.value = "";
  }
}

form.addEventListener("change", (event) => {
  syncOtherFields();
  if (event.target.value === "__other" && event.target.checked) {
    form.querySelector(`[data-other="${event.target.closest("[data-group]").dataset.group}"]`).focus();
  }
});
form.addEventListener("reset", () => setTimeout(syncOtherFields));

function buildEntry() {
  const data = new FormData(form);
  const behaviours = checked("behaviours");
  const duration = checked("duration")[0];
  const when = data.get("occurredAt");
  return {
    id: crypto.randomUUID(),
    occurredAt: (when ? new Date(when) : new Date()).toISOString(),
    createdAt: new Date().toISOString(),
    whatHappened: behaviours.length ? behaviours.join(", ") : "Big feelings",
    behaviours,
    triggers: checked("triggers"),
    durationMinutes: duration ? Number(duration) : null,
    durationLabel: duration ? DURATION_LABELS[duration] : "",
    helped: checked("helped"),
    who: checked("who")[0] || "",
    note: String(data.get("note") || "").trim(),
  };
}

async function appendEntry(entry) {
  // If-Match makes a concurrent save from the other phone fail instead of overwriting it.
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const latest = await loadLog();
    const merged = [...latest.entries.filter((existing) => existing.id !== entry.id), entry];
    try {
      await graph(`${itemPath()}/content`, {
        method: "PUT",
        headers: { "Content-Type": "application/json", "If-Match": latest.eTag },
        body: serialize(merged),
      });
      return merged;
    } catch (error) {
      if (error.status !== 412) throw error;
    }
  }
  throw new Error("busy");
}

async function save(event) {
  event.preventDefault();
  saveButton.disabled = true;
  say(formMessage, "Saving…");
  try {
    entries = await appendEntry(buildEntry());
    render();
    form.reset();
    setNow();
    say(formMessage, "Saved. You're doing a hard thing well.");
    say(storageStatus, `${entries.length} saved in your OneDrive log.`);
  } catch (error) {
    say(formMessage, "Not saved yet. Your answers are still here, so tap Save again.", true);
  } finally {
    saveButton.disabled = false;
  }
}

function setNow() {
  const now = new Date();
  $("#occurred-at").value = new Date(now.getTime() - now.getTimezoneOffset() * 60_000).toISOString().slice(0, 16);
}

async function shareLog() {
  const out = $("#share-out");
  try {
    const result = await (
      await graph(`${itemPath()}/createLink`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: "edit", scope: "anonymous" }),
      })
    ).json();
    const link = result.link.webUrl;
    try {
      await navigator.clipboard.writeText(link);
      out.textContent = "Link copied. Send it to your spouse (keep it private).";
    } catch {
      out.textContent = link;
    }
  } catch {
    out.textContent = "Couldn't create a link. Share the file from OneDrive instead.";
  }
}

$("#sign-in").addEventListener("click", () => msal.loginRedirect({ scopes: SCOPES }));
$("#sign-out").addEventListener("click", () => {
  localStorage.removeItem(REF_KEY);
  msal.logoutRedirect({ account });
});
$("#create-log").addEventListener("click", createLog);
$("#link-form").addEventListener("submit", connectByLink);
$("#refresh").addEventListener("click", refresh);
$("#share-log").addEventListener("click", shareLog);
form.addEventListener("submit", save);

setNow();
initAuth().catch((error) => {
  console.error(error);
  Object.keys(localStorage).filter((k) => k.startsWith("msal.")).forEach((k) => localStorage.removeItem(k));
  history.replaceState(null, "", location.pathname);
  $("#gate-text").textContent = `Sign-in didn't finish (${error.errorCode || error.message}). Tap the button to try again.`;
  show("gate");
});
