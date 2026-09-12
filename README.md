# Cascade Bruins Volleyball — Dinner Sign-Up (2026-27)

Single-page volunteer sign-up for home game dinners. Static frontend (Vercel) + Google Apps Script/Sheets backend.

> **Template:** Based on [claims-coach/CHS-Volleyball](https://github.com/claims-coach/CHS-Volleyball) (`chs-volleyball-signup-repo/`)

---

## Quick Start

### 1. Apps Script (backend)

1. Create a new Google Sheet (or use an existing one)
2. Go to **Extensions → Apps Script**
3. Delete any default code and paste the contents of `backend/Code.gs`
4. Replace the placeholder emails at the top:
   ```js
   const ORGANIZER_EMAIL = 'your-email@example.com';  // TODO: your email
   const COACH_EMAIL = 'coach@school.org';            // TODO: coach email
   ```
5. Click the `appsscript.json` file in the sidebar (enable "Show manifest file" in Project Settings if hidden) and replace with contents of `backend/appsscript.json`
6. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (or your Google Workspace domain)
7. Copy the **Web App URL** (looks like `https://script.google.com/macros/s/.../exec`)

### 2. Frontend

1. Open `index.html` and replace the placeholder:
   ```js
   const GAS_URL = "YOUR_GAS_WEB_APP_URL";
   ```
   with your actual Web App URL from step 1.

### 3. Deploy to Vercel

1. Push this repo to GitHub
2. In Vercel: **Add New → Project → Import** this repo
3. Leave build settings empty (it's a static site)
4. Deploy — your site is live!

---

## 2026-27 Home Game Dates

| Date | Opponent |
|------|----------|
| Tue Sept 15, 2026 | HOME vs Shorecrest |
| Wed Sept 16, 2026 | HOME vs Marysville Getchell |
| Thu Oct 8, 2026 | HOME vs Snohomish |
| Tue Oct 13, 2026 | HOME vs Stanwood |
| Mon Oct 19, 2026 | HOME vs Monroe |
| Wed Oct 21, 2026 | HOME vs Marysville-Pilchuck |
| Tue Oct 27, 2026 | HOME vs Everett |

*Source: MaxPreps Cascade (Everett, WA) 2026-27 schedule*

---

## Tournaments (TBD)

Tournament dates are not yet confirmed. When available, add them to the `GAMES` array in `index.html`:

```js
const GAMES = [
  // ... existing home games ...
  { id:"2026-MM-DD", label:"Sat Month DD — TOURNAMENT @ Location" }
];
```

Each entry needs:
- `id`: Date in `YYYY-MM-DD` format (used as unique key in the Sheet)
- `label`: Display text shown to volunteers

---

## Files

```
├── index.html              # Frontend (single-page app)
├── assets/
│   └── chs-logo.jpeg       # Cascade Bruins logo
├── backend/
│   ├── Code.gs             # Google Apps Script backend
│   └── appsscript.json     # Apps Script manifest
└── README.md
```

---

## Features

- **Single-slot lock**: One volunteer per date (enforced server-side)
- **Confirmation emails**: Sent to volunteer, coach, and organizer
- **Real-time roster**: Shows current signups
- **Drinks reminder**: Water/sports drinks only — no energy drinks or caffeine

---

## Notes

- No secrets are stored in this repo
- The `GAS_URL` placeholder must be replaced before the site will work
- For changes to the Apps Script, redeploy as a **new version** to update the live URL
