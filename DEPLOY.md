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
(Right now it points to a local test file, `steeltrack_local.db`, which already
holds the data from `master_database_final.xlsx`.)

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
2. **Data admin → Import master database** → upload `master_database_final.xlsx`
   → *Replace everything* → Import. (4,670 parts, 486 assemblies, 342.96 t.)
3. **Data admin → Project settings** → enter the project name.
4. **Users** → add your team. Access options:
   - no tokens = view only (Overview, Assemblies, Reports)
   - `Update Fit-Up`, `Update Welding`, `Update Painting`, `Update Site Delivery`
     = can book that stage
   - Full access = everything including Data admin / Users / Activity

## How progress works

- Progress is booked per **sub assembly**, in order:
  **Fit-up → Welding → Blasting & painting (D.O. No.) → Send to site (D.O. No.)**.
  A part that doesn't need painting goes straight from welding to site.
- A stage's kg = the part weight once that stage has a date (same as the Excel).
- **Update progress**: pick the stage, filter, tick sub-assemblies (or *Select all
  shown*), set the date (and D.O.), then Confirm. Mistakes: use *Undo* at the bottom.
- **Reports → Export master database** gives back the same 25-column Excel.
- Re-importing a newer Excel with *Add new parts & update details* keeps every
  date already booked in the app. A backup copy is made before every import
  (Data admin → Backups).

## Updating the app later

`git push` → Streamlit Cloud redeploys by itself.
