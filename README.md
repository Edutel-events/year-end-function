# Edutel Year-End Function 2026 · RSVP

RSVP page for the Edutel staff year-end function.

**Live:** https://edutel-events.github.io/year-end-function/

| Event | RSVP deadline |
|---|---|
| Friday, 20 November 2026 | Friday, 30 October 2026 |

## Features

- Collects name, work email, cellphone (optional), attendance and dietary requirement, with allergy details where needed
- Submitting again with the same email updates the existing response instead of duplicating it
- Responses are stored in a Google Sheet, with a Summary tab for catering numbers
- Passcode-protected organiser view: totals, dietary breakdown, all responses, copy attendee emails
- Single static page: no build step, no dependencies

## How it works

```
index.html (GitHub Pages) ──POST──▶ Apps Script web app ──▶ Google Sheet
```

The page posts JSON as `text/plain` so the browser skips the CORS preflight, which Apps Script can't answer. The organiser passcode is kept in the script's properties and checked server-side.

## Structure

```
index.html            page, styles and client script
apps-script/Code.gs   Apps Script back end, bound to the RSVP Sheet
assets/               logo: edutel-logo.svg or edutel-logo.png
```

## Back end

1. In the RSVP Sheet, open **Extensions → Apps Script** and paste in `apps-script/Code.gs`.
2. Run `setup` once and authorise it. This creates the `RSVPs` and `Summary` tabs.
3. Add a `PASSCODE` script property under **Project Settings → Script properties**.
4. **Deploy → New deployment → Web app**, executing as *Me* with access for *Anyone*.
5. Set `API_URL` in `index.html` to the deployment URL.

After changing `Code.gs`, use **Manage deployments → Edit → New version** so the URL stays the same.

## Customising

- **Colours, fonts, logo size:** the CSS variables at the top of `index.html`.
- **Logo:** add `assets/edutel-logo.svg` (or `.png`). A light version suits the dark background. Without a logo, an "E" monogram is shown.
- **Dietary options:** `DIET_OPTIONS` in both `index.html` and `Code.gs`.

## Data

Responses include contact details and dietary information, which can indicate religion or health. Keep access to the Sheet limited to the organisers and remove the data after the event.
