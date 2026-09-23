# Send to Skylight — Outlook Add-in

Adds a **"Send to Skylight"** button to any Outlook appointment (read or compose view).
Clicking it opens a compose window pre-filled with an ICS-friendly email to your Skylight Magic Import address, and downloads the `.ics` file automatically.

Supports:
- One-time and recurring events (daily, weekly, monthly by date or weekday, yearly)
- All-day events
- Categories (mapped to the email subject prefix Skylight uses)
- Body text, location, timezone

---

## Recommended: Deploy to GitHub Pages (works on managed machines)

This approach hosts the add-in files on **GitHub's free, publicly trusted HTTPS**, bypassing
all self-signed certificate issues. **You only need to do this once.**

### Step 1 — Push the repo to GitHub

1. Create a new GitHub repository at https://github.com/new  
   (e.g. `SkylightExtensions`). Keep it **public** (GitHub Pages is free for public repos).
2. On your **Mac** (the one running npm), open a terminal in this project root and run:

   ```bash
   git init
   git add .
   git commit -m "Add Skylight Outlook add-in"
   git remote add origin https://github.com/YOUR-GITHUB-USERNAME/SkylightExtensions.git
   git push -u origin main
   ```

### Step 2 — Enable GitHub Pages

1. Go to your repo on GitHub → **Settings** → **Pages**
2. Under "Build and deployment" set:
   - **Source:** Deploy from a branch
   - **Branch:** `main`   **Folder:** `/ (root)`
3. Click **Save** — GitHub shows a green banner with your URL in ~2 minutes:
   `https://YOUR-GITHUB-USERNAME.github.io/SkylightExtensions/`

### Step 3 — Generate your manifest

On your Mac, run the helper script:

```bash
bash outlook-addin/scripts/setup-github-manifest.sh
```

It asks for your GitHub username and repo name, then writes
`outlook-addin/manifest.github.final.xml` with the correct URLs.

Verify the file looks right — every URL should start with
`https://YOUR-GITHUB-USERNAME.github.io/...`

### Step 4 — Sideload via Outlook Web (OWA)

This works even on managed machines — no admin rights required.

1. Open https://outlook.office.com (or https://outlook.office365.com) in a browser
2. Click the **Settings** gear (top right) → **View all Outlook settings**
3. Go to **Mail → Customize actions** → scroll to **Add-ins**
   — *or* go directly to:
   https://outlook.office.com/mail/options/general/manage-addins
4. Click **Upload a custom add-in** → **Add from file…**
5. Select `outlook-addin/manifest.github.final.xml` and click **Open**
6. The add-in appears as "Send to Skylight" in your installed add-ins list.

> **Note:** The add-in is now installed for your **Microsoft 365 account**.
> It will appear in both Outlook on the Web and the **desktop Outlook app** on any
> machine signed into the same account — no IT permissions needed.

### Step 5 — Test it

1. Open any appointment in Outlook (desktop or web)
2. Look for the **"Send to Skylight"** button in the appointment ribbon/toolbar
3. Click it — the task pane opens showing the appointment details
4. Enter your Skylight Magic Import email address and click **Send to Skylight**

---

## Local Development (on your Mac)

### Run the local HTTPS dev server

```bash
cd outlook-addin
npm install
npm start
```

This starts a local HTTPS server on `https://0.0.0.0:3000` and prints your LAN IP.
The server generates self-signed certificates automatically on first run.

### Trust the certificate (Mac only — required for localhost testing)

```bash
sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain \
  -p ssl network-certs/ca.crt
```

Then use `manifest.network.xml` (which has been pre-filled with `https://YOUR-LAN-IP:3000/...`)
to sideload via OWA as described above.

### Validate the manifest

```bash
npm run validate            # validates manifest.xml (localhost)
npm run validate:network    # validates manifest.network.xml
```

### Browser testing (no Outlook needed)

Open `https://localhost:3000/app.html` directly in a browser.
The add-in detects it's running outside Outlook and pre-fills sample appointment data.

---

## Project structure

```
outlook-addin/
├── src/                        # Source files (served by local dev server)
│   ├── app.html
│   ├── app.css
│   ├── app.js                  # Main task pane logic
│   ├── ics.js                  # ICS file builder
│   └── timezones.js            # IANA timezone data
├── docs/                       # GitHub Pages deployment (copy of src/)
│   ├── app.html
│   ├── app.css
│   ├── app.js
│   ├── ics.js
│   ├── timezones.js
│   └── assets/                 # Icons
├── assets/                     # Icons (used by local dev server)
├── scripts/
│   ├── network-server.js       # Local HTTPS server with cert generation
│   └── setup-github-manifest.sh  # Fills in GitHub URLs in manifest template
├── manifest.xml                # Localhost manifest (local dev)
├── manifest.network.xml        # LAN IP manifest (cross-machine local testing)
├── manifest.github.xml         # GitHub Pages manifest TEMPLATE (fill in before use)
└── manifest.github.final.xml   # Generated by setup-github-manifest.sh ← USE THIS
```

---

## Keeping `docs/` in sync with `src/`

When you edit files in `src/`, copy them to `docs/` too:

```bash
cp outlook-addin/src/app.js outlook-addin/docs/app.js
cp outlook-addin/src/app.css outlook-addin/docs/app.css
cp outlook-addin/src/app.html outlook-addin/docs/app.html
cp outlook-addin/src/ics.js outlook-addin/docs/ics.js
cp outlook-addin/src/timezones.js outlook-addin/docs/timezones.js
```

Or add a one-liner to package.json `scripts` if you prefer.