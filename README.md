# BOTEC Calculator

A web application for creating back-of-the-envelope (BOTEC) cost-per-beneficiary estimates, modelled on the Punjab PDS and Rajasthan TAPF templates.

## Features

- **User authentication** via Supabase (email/password)
- **Save & load** BOTEC documents to/from a database
- **3-tab calculator**: beneficiary units, cost items, results
- **Correct formulae**: product chain for beneficiaries, managerial multiplier, buffer, CPB, CPB-excluding-logistics
- **Export**: CSV, Excel (.xlsx), JSON
- **Dashboard**: list, duplicate, delete documents

---

## Tech stack

| Layer | Tool |
|-------|------|
| Frontend | Vanilla HTML/CSS/JS (no build step) |
| Database + Auth | [Supabase](https://supabase.com) |
| Hosting | [Render](https://render.com) (static site) |
| Version control | GitHub |

---

## Setup guide

### 1. Fork / clone this repo

```bash
git clone https://github.com/YOUR_USERNAME/botec-calculator.git
cd botec-calculator
```

### 2. Create a Supabase project

1. Go to [supabase.com](https://supabase.com) → **New project**
2. Choose a name, database password, and region
3. Once created, go to **SQL Editor** → paste the contents of `sql/schema.sql` → **Run**
4. Go to **Project Settings → API** and copy:
   - **Project URL** (looks like `https://xxxx.supabase.co`)
   - **anon / public key**

### 3. Add your Supabase credentials

Open `js/supabase-client.js` and replace the placeholders:

```js
const SUPABASE_URL = 'https://xxxx.supabase.co';      // ← your Project URL
const SUPABASE_ANON_KEY = 'eyJhbGc...';               // ← your anon key
```

> **Security note:** The anon key is safe to expose in a static frontend — Supabase Row Level Security (RLS) ensures each user can only access their own data.

### 4. Enable Email auth in Supabase

Go to **Authentication → Providers → Email** — confirm it is enabled. Optionally disable "Confirm email" for easier local testing.

### 5. Push to GitHub

```bash
git add .
git commit -m "Initial BOTEC app"
git push origin main
```

### 6. Deploy to Render

1. Go to [render.com](https://render.com) → **New → Static Site**
2. Connect your GitHub account and select the `botec-calculator` repo
3. Settings:
   - **Build command:** leave blank (or `echo "no build"`)
   - **Publish directory:** `.`
4. Click **Create Static Site**
5. Render will deploy and give you a URL like `https://botec-calculator.onrender.com`

### 7. Update Supabase redirect URL (optional but recommended)

In Supabase → **Authentication → URL Configuration**:
- Add your Render URL to **Redirect URLs**: `https://botec-calculator.onrender.com/**`

---

## Local development

Since there is no build step, just open the files directly in a browser or use any static server:

```bash
# Python
python3 -m http.server 3000

# Node
npx serve .
```

Then open `http://localhost:3000`.

---

## Project structure

```
botec-calculator/
├── index.html           # Dashboard (list of saved BOTECs)
├── login.html           # Auth page (sign in / sign up)
├── calculator.html      # The BOTEC calculator
├── css/
│   └── style.css        # All styles
├── js/
│   ├── supabase-client.js   # Supabase init — add your keys here
│   └── calculator.js        # All calculator logic, save/load, exports
├── sql/
│   └── schema.sql           # Run once in Supabase SQL editor
├── render.yaml              # Render deployment config
└── README.md
```

---

## Calculation formulae

| Formula | Rule |
|---------|------|
| **Beneficiaries [Yr]** | Product of all unit rows for that year |
| **Row cost [Yr]** | Cost/unit × Units [Yr] |
| **Cost head total [Yr]** | Sum of all rows with that head |
| **Managerial multiplier [Yr]** | Internal Consulting [Yr] × multiplier % |
| **Sub-total [Yr]** | Sum(all cost heads) + managerial multiplier |
| **Buffer [Yr]** | Sub-total [Yr] × buffer % |
| **Total cost [Yr]** | Sub-total [Yr] + Buffer [Yr] |
| **CPB [Yr]** | Total cost [Yr] ÷ Beneficiaries [Yr] |
| **CPB (avg)** | Total 3-year cost ÷ Total beneficiaries (not average of averages) |
| **CPB excl. logistics [Yr]** | (Non-logistics sub-total + mgr multiplier) × (1 + buffer%) ÷ Beneficiaries [Yr] |

---

## Customisation

- **Add cost head presets:** edit the `PRESETS` object in `js/calculator.js`
- **Change default buffer/multiplier:** edit the HTML `value` attributes on `#bufferPct` / `#mgrMult`
- **Add more currencies:** extend the `<select id="currency">` in `calculator.html`
