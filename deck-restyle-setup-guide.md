# Setting Up the Deck Restyle Automation

You're on **n8n Cloud** (the trial banner and execution counter confirm it), not a self-hosted instance. That means `deck-tools-service` can't live on a private Docker network next to n8n — it needs its own public HTTPS URL that n8n Cloud can call over the internet. This guide uses **Railway** for that, since it builds straight from the Dockerfile you already have and gives you a public URL automatically.

Do the steps in order; later steps depend on earlier ones.

---

## Part 1: Deploy deck-tools-service on Railway

### 1.1 Generate a shared secret

The service now checks every request (except `/health`) for a header, since it's reachable from the public internet. Generate one:

```bash
openssl rand -hex 32
```

Save that string somewhere — you'll enter it in two places below.

### 1.2 Push the service to GitHub

Put these files (already in your downloads from this conversation) into a folder and push it as a GitHub repo — it can be private, and it can be the only thing in the repo:

```
deck-tools-service/
├── server.js
├── build_deck.js
├── Dockerfile
└── package.json
```

### 1.3 Deploy on Railway

1. Go to railway.com and sign in (GitHub login is easiest).
2. **New Project → Deploy from GitHub repo** → select the repo from 1.2. Railway detects the `Dockerfile` and builds automatically.
3. Once the service exists, open it → **Variables** tab → add:
   - `DECK_TOOLS_SECRET` = the string from 1.1
4. Open **Settings → Networking → Generate Domain**. You'll get a URL like:
   ```
   https://deck-tools-production-xxxx.up.railway.app
   ```
5. Wait for the deploy to finish (watch the **Deployments** tab — first build takes a few minutes since it installs LibreOffice).

### 1.4 Confirm it's live

```bash
curl https://deck-tools-production-xxxx.up.railway.app/health
```

Expected: `{"ok":true}`

If that fails, check the **Deployments → Logs** tab on Railway for a crash on startup.

---

## Part 2: Point both n8n workflows at the service

This step is already done for you in both workflows — I added a `deckToolsSecret` field and the header that sends it. You just need to fill in the two placeholder values.

### 2.1 For each workflow (Extract & Save Design System, and Restyle a Deck):

1. Open the workflow in n8n.
2. Click the **Deck Tools Config** node.
3. Set `deckToolsUrl` to your Railway URL from 1.3 (no trailing slash).
4. Set `deckToolsSecret` to the same string you put in `DECK_TOOLS_SECRET` on Railway (1.1).
5. Save the node, then save the workflow.

Do this in **both** workflows — they each have their own copy of this node.

---

## Part 3: Create the Anthropic API credential

Both workflows call Claude directly via HTTP, using an **HTTP Header Auth** credential.

### 3.1 Get an Anthropic API key

If you don't already have one: console.anthropic.com → API Keys → Create Key.

### 3.2 Create the credential in n8n

1. In either workflow, click the **Call Claude - Extract Design Spec** node (or **Call Claude - Map Target Content** in the Restyle workflow).
2. Under **Credential**, click **Create New**.
3. Choose credential type **Header Auth**.
4. Fill in:
   - **Name**: anything memorable, e.g. `Anthropic API Key`
   - **Header Name**: `x-api-key`
   - **Header Value**: your Anthropic API key
5. Save the credential.

### 3.3 Reuse it on the other Claude node

The other HTTP Request node calling `api.anthropic.com` (in the same or other workflow) needs the **same** credential selected — open it and pick the credential you just created from the dropdown, rather than creating a second one.

---

## Part 4: Run the extraction workflow (once)

This is the one-time step for a given design system.

1. Open **Extract & Save Design System** from your Personal workflows list.
2. Click the **Upload Reference Deck** node (the first one, the Form Trigger) to open its panel.
3. At the top of that panel you'll see a **Test URL / Production URL** toggle. Leave it on **Test URL** for now — Production URL only works once the workflow is activated, and you don't need that yet.
4. Close the node panel, then click **Execute workflow** at the bottom of the canvas (or **Test workflow**, depending on what n8n labels it). This arms the trigger and n8n opens the form in a new browser tab automatically.
5. In that new tab, fill in:
   - **Reference Deck (.pptx)** — the deck that's already in your brand design system.
   - **Design name** — something memorable and reusable, e.g. `SH4008 brand system`. Leave blank and it saves under `default`.
   - **Module / week identifier** — optional, used for the breadcrumb/footer template, e.g. `SH4008 Culture, Society & Ethics - Week 4`.
6. Submit. **Keep that tab open** — this is one continuous session, not a fire-and-forget submission. It'll sit on a loading state while it converts slides to images and calls Claude.
7. The same tab will then show a **review page** with the extracted design spec as JSON. Check it over:
   - Hex colours under `colors` match your brand.
   - `persistentElements.logo.path` points to a sensible filename.
   - `layouts` has an entry for each distinct slide type you expect.
   - Edit the JSON directly in the box if anything's wrong.
8. Submit the review form. The same tab shows a confirmation that the design name was saved.

You can double check it landed correctly: in n8n, go to **Personal → Data tables → design_specs**, and you should see a row with your design name.

That design name now lives in the data table. You will not need to repeat this step unless the brand design system itself changes.

---

## Part 5: Run the restyle workflow (every time you have a new deck)

1. Open **Restyle a Deck** from your Personal workflows list.
2. Click **Execute workflow** at the bottom of the canvas. n8n opens the form in a new tab.
3. Upload the **Target Deck (.pptx)** you want restyled.
4. Enter the **exact same design name** you used in Part 4 (e.g. `SH4008 brand system`). A mismatch gives you a clear "no saved design system found" message instead of a broken build.
5. Submit and keep the tab open — it's converting the target deck, calling Claude, and building the file.
6. When it finishes, the restyled `.pptx` downloads directly in that tab.

Repeat Part 5 as many times as you like, for as many target decks as you like, without ever touching Part 4 again.

### Once you're past testing

Test workflow / Execute workflow only works while you're at the n8n editor watching. For everyday use without opening the editor each time:
1. Open each workflow and toggle it **Active** (top right).
2. Open the Form Trigger node, switch the toggle to **Production URL**, and copy that link.
3. Bookmark the production URL for whichever workflow you'll reuse (almost always **Restyle a Deck**) — that's now a standalone link you can open directly any time, no need to open the editor.

---

## Quick troubleshooting

| Symptom | Likely cause |
|---|---|
| Workflow fails immediately on the HTTP nodes to `deckToolsUrl` | Wrong URL, Railway service still building, or you forgot `https://` |
| HTTP nodes to deck-tools fail with a 401 | `deckToolsSecret` in the Deck Tools Config node doesn't match `DECK_TOOLS_SECRET` on Railway |
| "No saved design system found" on the Restyle workflow | Design name typo, or you haven't run Part 4 yet for that name |
| Claude call fails with a 401 | Anthropic credential not set, or set on only one of the two Claude HTTP nodes |
| Restyled deck looks visually off | The design spec needs editing — rerun Part 4 and correct the JSON on the review page before saving |
| Restyled deck is missing slide content | Check the target deck isn't unusually large (very long decks may hit Claude's output token limit on the content-mapping call) |
