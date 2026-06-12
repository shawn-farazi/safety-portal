# Advanced Ironworks Safety Portal — Setup Guide

This is the full web app: a dashboard plus four iPad forms (Daily Safety Meeting,
Hazard Assessment, Accident Investigation, New Employee Orientation). Each form
builds a PDF on the iPad and, once connected, **saves it straight to OneDrive with
one tap — no folder picking.**

## What's in this folder

| File | Purpose |
|---|---|
| `index.html` | Dashboard / launcher. Sign in to OneDrive here once. |
| `Daily Safety Meeting.html` | Daily meeting + sign-in sheet. |
| `Hazard Assessment Form.html` | Quarterly PPE hazard assessment. |
| `Accident.html` | Accident / incident investigation. |
| `New Employee Orientation.html` | New-hire orientation checklist. |
| `aiw-config.js` | **The only file you edit.** Holds the two IDs from the app registration. |
| `aiw-graph.js` | Sign-in + OneDrive upload engine. Do not edit. |
| `aiw-save.js` | Decides: upload to OneDrive, or fall back to the Share Sheet. Do not edit. |
| `blank.html` | Tiny page Microsoft sign-in returns to. Do not edit. |

## How the saving works

When you tap **Save PDF**:

1. If OneDrive is connected → the PDF uploads to
   `OneDrive / Safety Records / <form folder>` and you get a green
   "Saved to OneDrive" confirmation. The four folders are created automatically:
   - `Safety Records / Daily Safety Meetings`
   - `Safety Records / Hazard Assessments`
   - `Safety Records / Accident Reports`
   - `Safety Records / New Employee Orientation`
2. If OneDrive is **not** set up yet → it falls back to the old iPad Share Sheet,
   so nothing ever breaks while setup is pending.

You sign in **once** on the dashboard; all four forms reuse that sign-in.

---

## One-time setup (the part we'll do live together)

Two things must exist before one-tap save works. Both are quick.

### Step 1 — Register the app in Microsoft Entra
1. Go to **entra.microsoft.com** → **Applications → App registrations → New registration**.
2. Name: `AIW Safety Portal`.
3. Supported account types: **Accounts in this organizational directory only (single tenant)**.
4. Platform → **Single-page application (SPA)**, Redirect URI: your hosting URL + `/blank.html`
   (e.g. `https://aiw-safety.example.com/blank.html`). We set this after Step 2.
5. **Register.**
6. On the **Overview** page, copy **Application (client) ID** and **Directory (tenant) ID**.
7. **API permissions** → Add → Microsoft Graph → **Delegated** → add **Files.ReadWrite**
   (User.Read is already there). Click **Grant admin consent** so staff aren't prompted.

### Step 2 — Host the app so sign-in works
Microsoft sign-in only runs from a real `https://` web address, not from a file
opened in the Files app. So the folder needs to live on a small static host. Free
options we can use:
- **Azure Static Web Apps** (free tier, inside your Microsoft account), or
- **GitHub Pages** (free).

Whichever we pick gives a URL like `https://something/`. That becomes:
- the address you open on the iPad, and
- the redirect URI in Step 1 (that URL + `blank.html`).

### Step 3 — Plug in the two IDs
Open `aiw-config.js` and replace:
- `PASTE_CLIENT_ID_HERE` → Application (client) ID
- `PASTE_TENANT_ID_HERE` → Directory (tenant) ID

Re-upload `aiw-config.js` to the host. Done.

### Step 4 — On the iPad
1. Open the hosted `index.html`.
2. Tap **Sign in**, use your Advanced Ironworks account. The dot turns green.
3. (Optional) Add to Home Screen for an app-like icon.
4. Open any form, fill it out, tap **Save PDF** → it lands in OneDrive.

---

## Notes
- Files save to the **signed-in user's** OneDrive for Business. If you'd rather they
  land in a shared SharePoint library later, only `aiw-graph.js` needs a small change.
- The iPad stays signed in; staff won't re-enter a password every day.
- Filenames already include the form type and date, e.g.
  `Safety_Meeting_Hearing_Protection_2026-06-12.pdf`.
