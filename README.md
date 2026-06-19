# deck-tools sidecar service

A small standalone Node service that does the two things n8n itself can't
do any more: convert a PPTX to slide images (LibreOffice + poppler-utils),
and build a restyled PPTX with pptxgenjs. It runs in its own container,
next to your n8n container, and n8n talks to it over HTTP.

## Why a separate container

n8n's official Docker image went distroless from v2.x onward, there's no
`apk`/`apt-get`, no shell, and no way to add LibreOffice to it directly
any more (this used to work in v1.x, but isn't reliable going forward).
Rather than depend on which n8n image you're running, this service lives
on its own, built from a normal `node:22-slim` base where `apt-get` still
works. n8n calls it with the HTTP Request node, exactly like calling any
other API.

## Security: shared-secret header

Since this service may be reachable from the public internet (not a private
network n8n already shares), it checks every request except `/health` for a
header:

```
x-deck-tools-secret: <a long random string you choose>
```

Set it via the `DECK_TOOLS_SECRET` environment variable when running the
container. If you don't set it, auth is disabled — fine for purely local
testing, not for anything internet-reachable. Generate a good secret with:

```bash
openssl rand -hex 32
```

## Deploy it on Railway

Railway builds straight from the Dockerfile here and gives you a public
HTTPS URL with no server/TLS management.

1. Push this `deck-tools-service` folder to a GitHub repo (it can be a
   private repo, and it can be the only thing in the repo).
2. In Railway: **New Project → Deploy from GitHub repo** → pick that repo.
   Railway detects the `Dockerfile` and builds automatically.
3. Open the service → **Variables** tab → add:
   - `DECK_TOOLS_SECRET` = the random string you generated above
4. Open the service → **Settings → Networking → Generate Domain**. Railway
   gives you a URL like `https://deck-tools-production-xxxx.up.railway.app`.
5. Test it:
   ```bash
   curl https://deck-tools-production-xxxx.up.railway.app/health
   ```
   Expect `{"ok":true}`.

That URL (no trailing slash) is your `deckToolsUrl` for the n8n workflows.
Every HTTP Request node calling this service also needs the
`x-deck-tools-secret` header set to the same value as `DECK_TOOLS_SECRET`.

### Alternative: your own server + docker-compose

If you'd rather run this on a VPS or alongside a self-hosted n8n on the same
docker-compose network, the original approach still works:

```yaml
services:
  n8n:
    # ... your existing n8n service, unchanged ...

  deck-tools:
    build: ./deck-tools-service
    restart: unless-stopped
    environment:
      - DECK_TOOLS_SECRET=your-random-secret-here
    expose:
      - "4000"
```

```bash
docker compose up -d --build deck-tools
```

If n8n and this service are on the same private docker-compose network,
`deckToolsUrl` is `http://deck-tools:4000` and you technically don't need
the secret (nothing outside the network can reach it) — but setting one
costs nothing and protects you if you ever expose the port.

## API

Every request below except `/health` must include the
`x-deck-tools-secret` header described above (if you set one).

### `POST /pptx-to-images`

Raw binary body = the `.pptx` bytes (no multipart wrapper).

```json
{ "count": 5, "images": [{ "filename": "slide-1.jpg", "base64": "..." }] }
```

In n8n: HTTP Request node, **Body Content Type: "Binary File"** (`contentType: 'binaryData'`),
`inputDataFieldName` = the binary property holding the PPTX. The response
is JSON, one item per image; use a **Convert to File** node (operation
`toBinary`, source property `base64`) on each `images[i]` to turn it into
a binary image n8n can pass to the Anthropic node.

### `POST /pptx-extract-assets`

Same raw-binary input as above. Returns the deck's embedded pictures
(logo, screenshots, etc. from `ppt/media/`):

```json
{ "count": 2, "assets": [{ "filename": "image-1-1.png", "base64": "..." }] }
```

Used to pull out the reference deck's logo and the target deck's
embedded images, so they can be passed straight through to `/build-deck`
without ever being re-encoded or summarised away.

### `POST /build-deck`

JSON body:

```json
{
  "designSpec": { "...": "see design_spec schema below" },
  "slideContent": [ { "...": "see slide_content schema below" } ],
  "assets": [{ "filename": "image6.png", "base64": "..." }]
}
```

Returns the built `.pptx` as raw binary
(`application/vnd.openxmlformats-officedocument.presentationml.presentation`).
In n8n: HTTP Request node with **Response Format: File** captures this
directly as a binary item, no extra conversion needed.

## JSON schemas

These are the two documents the Claude steps in the n8n workflow need to
produce. `example_design_spec.json` and `example_slide_content.json` in
this folder are filled-in, working examples (using the SH4008 reference
deck's actual values) you can use as few-shot examples in the Claude
prompts.

### `design_spec.json`

```
{
  "canvas":   { "width": 10, "height": 7.5 },
  "colors":   { "steelBlue": "4472C4", "paleBlue": "DCE6F1", "navy": "1F3864",
                "white": "FFFFFF", "black": "000000", "green": "70AD47",
                "paleGreen": "E2EFDA", "orange": "C55A11", "amber": "FFC000",
                "grey": "595959" },
  "typography": { "fontFace": "Calibri" },
  "persistentElements": {
    "breadcrumb":  { "x", "y", "w", "h", "size", "color", "text" },
    "logo":        { "x", "y", "w", "h", "path" or "data" },
    "sectionTitle":{ "x", "y", "w", "h", "size", "color" },
    "subtitle":    { "x", "y", "w", "h", "size", "color" },
    "footerBar":   { "x", "y", "w", "h", "color" },
    "footerText":  { "x", "y", "w", "h", "size", "color", "text" }
  },
  "layouts": { "A": {...}, "B": {...}, ..., "K": {...} }
}
```

Anything left out of `layouts` or `persistentElements` falls back to the
SH4008 reference measurements baked into `build_deck.js` as defaults, so
a partial extraction still renders something reasonable. The full set of
default measurements for each layout type is at the top of
`build_deck.js`, that file is the authoritative reference for what each
layout type's config object accepts.

### `slide_content.json`

An array, one entry per output slide:

```json
{
  "layoutType": "A",
  "sectionTitle": "...",
  "subtitle": "...",
  "breadcrumb": "... (optional override)",
  "footerText": "... (optional override)",
  "notes": "... (optional speaker notes)",
  "content": { "...": "shape depends on layoutType, see below" }
}
```

`content` by layout type:

| Type | Meaning | `content` shape |
|---|---|---|
| A | Two-column setup cards | `{ cards: [{header, body}, {header, body}] }` |
| B | Title / hero slide | `{ moduleLabel, headline, tagline, lecturer }` |
| C | Numbered question grid | `{ cards: [{number, text}, ...] }` |
| D | Full-width image | `{ image: "relative/path/in/assets.png" }` |
| E | Numbered step grid | `{ startPoint: {label, context}, steps: [{number, text}, ...] }` |
| F | Flowchart / process map | `{ startPoint: {label, context}, stageLabel, cards: [{header, body}], annotation }` |
| G | Table / framework grid | `{ headers: [...], rows: [[...], ...] }` |
| H | 3-column weighting panel | `{ columns: [{header, subtitle, body}, x3] }` |
| I | Annotated text with pills | `{ citationHeader, body, pills: [{label, color}] }` |
| J | Referencing / resource slide | `{ image: "..." }` OR `{ links: ["...", ...] }` |
| K | Assessment / module info | `{ headline, infoBoxes: ["...", ...], cards: [{number, text}, x4] }` |

Images referenced by relative path (Type D, J, and the logo) must also
appear in the request's top-level `assets` array with matching
`filename`, so the service can write them to disk before pptxgenjs reads
them.

## Files

- `server.js` — the Express service (two endpoints above)
- `build_deck.js` — the generalized pptxgenjs renderer, run as a
  subprocess by `/build-deck`. Can also be run standalone:
  `node build_deck.js design_spec.json slide_content.json out.pptx [assetsDir]`
- `Dockerfile` — builds the service with LibreOffice + poppler-utils
- `example_design_spec.json`, `example_slide_content.json` — working
  examples based on the SH4008 reference deck

Both endpoints were tested locally end to end (pptx in -> images out,
spec+content in -> valid pptx out, opened and re-rendered to confirm)
before being handed off here.
