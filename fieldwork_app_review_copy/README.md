# Fieldwork Project Manager

This is a Flask web app for managing geotechnical fieldwork progress. It currently supports projects, boreholes, CPTU/SCPTU, test pits, geophysics, custom test methods, map imports, calendar planning, dashboard summaries, rates, and budgeting history.

## Current App Structure

- `app.py` - Flask backend, routes, calculations, SQLite/PostgreSQL database access, and Supabase login.
- `templates/` - Jinja HTML pages.
- `static/` - CSS, logo, and local Leaflet fallback assets.
- `data.db` - local SQLite database used when `DATABASE_URL` is not set.
- `supabase/migrations/001_initial_schema.sql` - Supabase PostgreSQL database setup.
- `.env.example` - environment variable template.
- `render.yaml` - Render deployment template.
- `SETUP_GUIDE.md` - step-by-step hosted setup checklist.
- `tools/check_setup.py` - local setup verification helper.

## Recommended Hosting Setup

Use this setup:

1. GitHub for the source code.
2. Render for hosting the Flask web app.
3. Supabase for PostgreSQL database storage and email/password login.

Render is recommended because this app already has a Flask backend. Vercel and Netlify are excellent for static sites, but this app is not frontend-only. Keeping Flask avoids a large rewrite.

For the easiest hosted setup, follow:

```text
SETUP_GUIDE.md
```

## Local Setup

To run locally with the existing SQLite database:

```powershell
python app.py
```

Then open:

```text
http://127.0.0.1:5000
```

If you create a `.env` file and leave `DATABASE_URL` blank, the app still uses `data.db`.

To check your current setup:

```powershell
python tools/check_setup.py
```

## Supabase Setup

1. Create a Supabase project.
2. Open the Supabase SQL Editor.
3. Copy and run the SQL from:

```text
supabase/migrations/001_initial_schema.sql
```

4. Go to Authentication, then Users.
5. Add your own user email and password.
6. Keep public signups disabled unless you deliberately want outside users to register themselves.

## Environment Variables

Copy `.env.example` to `.env` for local testing, or add the same values in Render.

Required for hosted use:

```text
DATABASE_URL=your Supabase PostgreSQL connection string
SUPABASE_URL=https://your-project-ref.supabase.co
SUPABASE_ANON_KEY=your Supabase anon public key
FLASK_SECRET_KEY=a long random secret value
LOCAL_AUTH_BYPASS=false
```

Notes:

- `DATABASE_URL` controls cloud data storage.
- `SUPABASE_URL` and `SUPABASE_ANON_KEY` turn login protection on.
- `FLASK_SECRET_KEY` keeps login sessions stable and private.
- `LOCAL_AUTH_BYPASS=true` should only be used for local testing.

## Deploy to Render

1. Push this folder to GitHub.
2. In Render, create a new Web Service.
3. Connect the GitHub repository.
4. Use these settings:

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
```

5. Add the environment variables listed above.
6. Deploy the service.
7. Open `https://your-render-app.onrender.com/setup-status`.
8. Confirm `ok`, `database_connected`, and `auth_configured` are all `true`.
9. Open the normal Render URL.
10. Log in with the Supabase Auth user you created.

## Moving Existing Local Data

Use the new JSON migration flow:

1. Open the local app.
2. Open a project.
3. Click `Export JSON`.
4. Open the hosted app.
5. On the home page, use `Import project JSON`.
6. Open the imported project and check the dashboard, map, calendar, and data tabs.

This avoids manually copying SQLite database tables.

## Mobile Testing Checklist

After deployment, test on your phone:

- Log in and log out.
- Open the project list.
- Open a project.
- Check Overview, Dashboard, Map, Calendar, Budgeting, and Data Management.
- Add or edit one fieldwork item.
- Refresh the page and confirm the change is still there.
- Export a project JSON file if needed.

## Important Behaviour

- Skipped items are excluded from totals, rates, projections, and budgeting calculations.
- Local SQLite still works for testing.
- Hosted PostgreSQL is used only when `DATABASE_URL` is set.
- Login protection is enabled only when Supabase Auth settings are set.

## Future Improvements

- Add project-level team permissions so different users can only see assigned projects.
- Add mobile photo uploads for field evidence.
- Add offline queueing for poor site connectivity.
- Add a daily progress email or Teams summary.
- Add role levels such as viewer, editor, and admin.
