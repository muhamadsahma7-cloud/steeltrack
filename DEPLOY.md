# SteelTrack — putting it online

Same setup as Team Piping: **Supabase** (online database) + **Streamlit Community
Cloud** (the website). Both are free.

The app creates its own tables the first time it connects, so you don't run any SQL.

## 1. Create the database (Supabase)

1. Go to <https://supabase.com/dashboard> → **New project** (e.g. `steeltrack`, region
   *Southeast Asia (Singapore)* or the same region as Team Piping).
   Use a **new** project so it doesn't clash with Team Piping's tables.
   (The free plan allows 2 active projects.)
2. Save the database password you set.
3. **Connect** (top bar) → **Session pooler** → copy the URI. It looks like
   `postgresql://postgres.abcdxyz:[YOUR-PASSWORD]@aws-0-ap-southeast-1.pooler.supabase.com:5432/postgres`
4. Turn it into the app's URL:
   - change `postgresql://` to `postgresql+psycopg2://`
   - put your password in place of `[YOUR-PASSWORD]`; write an `@` in the password as `%40`

## 2. Test it on your PC (optional)

Edit `streamlit_app/.streamlit/secrets.toml`:

```toml
[connections.steel]
url = "postgresql+psycopg2://postgres.abcdxyz:PASSWORD@aws-0-ap-southeast-1.pooler.supabase.com:5432/postgres"
```

Double-click `run.bat`. The browser opens on the login page.
(To test without Supabase, use `url = "sqlite:///steeltrack_local.db"`. That file
already holds the data from `MASTER DATABASE 20261015.xlsx`.)

## 3. Put the code on GitHub

```bash
cd "D:\my latest software\Steeltrack"
git init
git add .
git status          # check: NO secrets.toml, NO .db, NO .xlsx in the list
git commit -m "SteelTrack - steel structure progress"
git branch -M main
git remote add origin https://github.com/muhamadsahma7-cloud/steeltrack.git
git push -u origin main
```

(Create the empty `steeltrack` repo on github.com first.)

## 4. Create the website (Streamlit Cloud)

1. <https://share.streamlit.io> → **Create app** → *Deploy a public app from GitHub*.
2. Repository `muhamadsahma7-cloud/steeltrack`, branch `main`,
   **Main file path:** `streamlit_app/app.py`.
3. **Advanced settings → Secrets**, paste:
   ```toml
   [connections.steel]
   url = "postgresql+psycopg2://postgres.abcdxyz:PASSWORD@aws-0-ap-southeast-1.pooler.supabase.com:5432/postgres"
   ```
4. Deploy. Keep app **Sharing = public**. The app's own login protects it,
   just like Team Piping.

## 5. First sign-in

1. Sign in as **admin** / **admin123**. Change the password straight away
   (sidebar → Change password).
2. **Data admin → Import master database** → upload the latest master Excel (e.g. `MASTER DATABASE 20261015.xlsx`)
   (sheet **Progress**) → *Replace everything* → Import.
   (1,653 rows, 259 sub-assemblies, 725 pieces, NAEC scope 149.28 t.)
3. **Data admin → Project settings** → enter the project name.
4. **Users** → add your team. Access options:
   - no tokens = view only (Overview, Assemblies, Reports)
   - `Update Fit-Up`, `Update Welding`, `Update Painting`, `Update Site Delivery`
     = can book that stage
   - Full access = everything including Data admin / Users / Activity

## How progress works

- The Excel layout is the **Progress** sheet of `MASTER DATABASE 20261015.xlsx` (27 columns;
  older files without `Welded?` still import, as all welded).
  `WEIGHT` = one piece; row weight = `No.` × `WEIGHT`.
- `TOTAL WEIGHT` heads a block of rows: block weight × number of identical copies.
  So A5 (0.9 kg, TOTAL 8.1) is made **9 times**; B132-1…10 (TOTAL 287) once each.
  The import screen lists any `TOTAL WEIGHT` that doesn't divide evenly (TTR22 now).
- Progress is booked as **pieces per sub-assembly**:
  **Fit-up → Welding → Blasting & painting → Send to site (D.O. No.)**.
  Each stage can only take pieces the previous stage has finished.
- `Welded?` = **NO** (purlins, plain beams, angles) is never booked for welding: booking its
  fit-up fills in welding automatically (same pieces, same date).
  Routes: YES/No: Fit-up → Welding → Site · YES/Yes: Fit-up → Welding → Painting → Site ·
  NO/No: Fit-up → Site · NO/Yes: Fit-up → Painting → Site (Welded? / Painting status).
- `Painting status` = **Yes** must be painted before site; **No** skips blasting & painting.
- `NAEC Scope` = **NO** stays in the list but is not counted in progress or tonnage.
- A stage's kg = pieces done × weight of one copy.
- **Update progress**: pick the stage, filter, type pieces in the **Book** column
  (or tick *Book all available*), set the date (and D.O. for site), then Confirm.
  Mistakes: *Undo* at the bottom (remove the later stage's booking first).
- Sent-to-site dates in the Excel without fit-up/welding dates are taken as
  fit-up/welding done too (no date), so totals add up.
- **Reports → Export master database** gives back the same 27-column Excel.
- Re-importing a newer Excel with *Add new parts & update details* keeps every
  booking made in the app. A backup is made before every import
  (Data admin → Backups).

## Updating the app later

`git push` → Streamlit Cloud redeploys by itself.
