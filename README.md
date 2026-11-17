# Suds Construction Monitoring

Monitoring system for new laundry shop builds: supplier enrollment → invitation to quote → quotation &
negotiation → award (with regret notices) → downpayment → mobilization → weekly updates, milestones,
concerns, change orders → payment requests & proof of settlement → turnover. Includes a cost dashboard per project.

| Part | Where it lives |
|---|---|
| Page (what people see) | `public/` – published by **Cloudflare** from this GitHub repo |
| Backend (data, emails, uploads) | `apps-script/Code.gs` – **Google Apps Script** attached to a Google Sheet |
| Files (photos, videos, quotes, proofs) | A Google Drive folder made by the script |
| Owner access | Inside the **Odoo portal** (`odoo/portal-construction-view.xml`) |

**Who uses it**
- **Admin (SPFC)**: opens the Cloudflare link → *Admin* tab → admin PIN. Accredits suppliers, adds owners, sees all projects.
- **Owner (Odoo portal user)**: opens *My Account* in the Odoo portal; the app appears there (no extra login).
- **Supplier/contractor**: opens the Cloudflare link, registers, logs in with email + PIN.

---

## Setup (one time)

### 1. Google Sheet + Apps Script (backend)
1. Create a new Google Sheet named **Suds Construction Monitoring** (using the Suds Google account that should send the emails).
2. **Extensions → Apps Script**. Delete what's in `Code.gs` and paste the whole `apps-script/Code.gs` from this repo.
3. Click ⚙️ **Project Settings** → tick **Show "appsscript.json" manifest file**. Open `appsscript.json` in the editor and replace it with `apps-script/appsscript.json` from this repo.
4. Back in `Code.gs`, choose **setup** in the function dropdown → **Run** → allow the permissions.
5. Open **Execution log**: it shows the **ADMIN PIN**. Keep it. (You can change it any time in Project Settings → Script properties → `ADMIN_PIN`.)
6. **Deploy → New deployment → Web app**. Execute as: **Me**. Who has access: **Anyone**. Click **Deploy** and copy the **Web app URL** (ends with `/exec`).
7. Open `public/index.html` and paste the URL into `window.SUDS_API_URL = '...'`.

> When you change `Code.gs` later: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** (the URL stays the same).

### 2. Cloudflare (page)
1. Cloudflare dashboard → **Workers & Pages → Create → Import a repository** → choose **suds-proj/suds-construction**.
2. Leave the defaults (the repo already has `wrangler.jsonc`) → **Deploy**.
3. Copy your link (e.g. `https://suds-construction.<your-subdomain>.workers.dev`).
4. In Apps Script → Project Settings → Script properties set:
   - `APP_URL` = your Cloudflare link (used in supplier emails)
   - `OWNER_URL` = your Odoo portal link, e.g. `https://yourcompany.odoo.com/my` (used in owner emails)

Every push to GitHub redeploys automatically.

### 3. Odoo portal (owners)
1. **Studio** on Contacts → add a **Text** field labeled **Construction Key** (technical name `x_studio_construction_key`).
2. Create a contact tag **Construction Owner**.
3. Add the view in `odoo/portal-construction-view.xml` (steps are inside the file). Replace the Cloudflare link in it.
4. In the app (Admin → **Owners → + Add owner**), add the owner. Click **Copy** on their key and paste it into the owner's Odoo contact (**Construction Key**), and give the contact the **Construction Owner** tag.

The owner now sees *Construction Monitoring* on their portal home page.

### 4. Daily reminders
`setup()` already created a daily 8 AM trigger (`dailyJobs`):
- quote due tomorrow → reminder to supplier; past deadline → invitation **Expired**
- Friday → reminds contractors to post their weekly update
- Monday → tells the owner which packages had **no update in 7 days**

---

## Flow
1. Supplier **registers** → admin **accredits** (email sent).
2. Owner creates a **project** and **packages** (General Contractor, CCTV, WiFi, Sprinkler, Smoke Detector, LPG, Aircon, Glass, Signage…).
3. Owner **invites** accredited suppliers to quote (brief, deadline, attachments) → emails.
4. Supplier **submits a quote** (line items, days, payment terms, warranty) or **declines**.
5. Owner **accepts / negotiates (counter-offer) / declines** → emails. Supplier revises and can mark **Final quote** (locked).
6. Owner **awards** → winner gets an **award notice**; everyone else who quoted gets a **thank-you/regret notice** (winner and price not shown).
7. If a **downpayment** is required, owner records it with proof → supplier **confirms received**.
8. Supplier submits **mobilization** (date, days → target end, workers, foreman, materials).
9. Weekly: supplier posts **weekly update** (% done, work done, next plan, manpower, delays, ≥3 photos, video). Ticks **milestones** with photos.
10. Either side raises **concerns**; supplier requests **change orders** (owner approves → contract and end date update).
11. Supplier **requests payment** → owner **pays & uploads proof of settlement** → supplier confirms.
12. Supplier submits the **turnover checklist** → owner clicks **Turned Over** (warranty starts).

## Data (Google Sheet tabs)
Owners · Suppliers · Projects · Packages · Invitations · Quotes (every revision kept) · Payments · PayRequests ·
Mobilization · Milestones · Updates · Concerns · ChangeOrders · Turnover · EmailLog
