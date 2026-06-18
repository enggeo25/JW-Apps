# Hosted Setup Guide

Follow this in order. You only need to do most of these steps once.

## 1. Supabase Database

1. Open your Supabase project.
2. Go to `SQL Editor`.
3. Open this file in the repo:

```text
supabase/migrations/001_initial_schema.sql
```

4. Copy all the SQL.
5. Paste it into Supabase SQL Editor.
6. Click `Run`.

This creates the project, map item, import backup, and historical rate tables.

## 2. Supabase Login User

1. In Supabase, go to `Authentication`.
2. Go to `Users`.
3. Click `Add user`.
4. Add your own email and password.
5. Keep public signups disabled unless you want other people to create accounts themselves.

## 3. Supabase Values to Copy

You need three values.

### Supabase URL

Go to:

```text
Project Settings -> API -> Project URL
```

Copy that into:

```text
SUPABASE_URL
```

### Supabase anon key

Go to:

```text
Project Settings -> API -> Project API keys -> anon public
```

Copy that into:

```text
SUPABASE_ANON_KEY
```

### Database URL

Go to:

```text
Project Settings -> Database -> Connection string
```

Use the connection string for the pooler if Render has trouble with the direct connection.

Paste it into:

```text
DATABASE_URL
```

It should look roughly like this:

```text
postgresql://postgres.PROJECT_REF:PASSWORD@aws-0-REGION.pooler.supabase.com:6543/postgres
```

Replace the password placeholder with your real database password.

## 4. Render Setup

1. Open Render.
2. Create a new `Web Service`.
3. Choose your GitHub repository.
4. Use these settings:

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
Health Check Path: /healthz
```

5. Add these environment variables:

```text
DATABASE_URL=your Supabase database connection string
SUPABASE_URL=your Supabase project URL
SUPABASE_ANON_KEY=your Supabase anon public key
FLASK_SECRET_KEY=a long random secret
LOCAL_AUTH_BYPASS=false
```

For `FLASK_SECRET_KEY`, use any long random text. Do not reuse your Supabase password.

## 5. First Hosted Test

After Render deploys, open:

```text
https://your-render-app.onrender.com/setup-status
```

You want to see:

```text
"ok": true
"database_connected": true
"auth_configured": true
```

Then open the normal app URL and log in.

## 6. Move Local Data

1. Open the local copy of the app.
2. Open a project.
3. Click `Export JSON`.
4. Open the hosted app.
5. Log in.
6. Use `Import project JSON` on the home page.
7. Open the imported project and check the map, calendar, dashboard, budgeting, and data tabs.

## 7. Phone Test

On your phone:

1. Open the Render app URL.
2. Log in.
3. Open a project.
4. Edit one item status.
5. Refresh the page.
6. Confirm the edit is still there.

That confirms the phone is updating the cloud database.

## Common Fixes

If `/setup-status` says the database failed:

- Check `DATABASE_URL`.
- Make sure the database password is correct.
- Try the Supabase pooler connection string instead of the direct connection string.

If login does not work:

- Check `SUPABASE_URL`.
- Check `SUPABASE_ANON_KEY`.
- Confirm the user exists in Supabase Authentication.

If Render deploy fails:

- Check that `requirements.txt` is committed.
- Check that the start command is exactly `gunicorn app:app`.
- Check Render logs for the first red error line.
