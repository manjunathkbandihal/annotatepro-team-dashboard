# AnnotatePro — Online Storage Setup

The dashboard now supports Supabase online storage while keeping localStorage as a fallback.

## 1. Create a Supabase project

Create a project in Supabase and open **SQL Editor**.

Run the complete contents of:

`SUPABASE_SCHEMA.sql`

## 2. Add your environment variables

Copy:

`.env.example` → `.env`

Then replace:

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`

Use the project's browser-safe anon/publishable key. Do **not** put a Supabase service-role key in this Vite frontend.

## 3. Install dependencies

```bash
npm install
```

## 4. Start the dashboard

```bash
npm run dev
```

When Supabase is configured, the Settings page should show **Online**.

## 5. First online sync

The app first checks Supabase for dashboard data.

If the cloud table is empty, the current local dashboard data is uploaded automatically.

After that, changes such as:

- add team member
- edit team member
- delete team member
- add/edit/delete project
- issue changes
- Excel imports
- accuracy report imports

are saved to Supabase.

## Important

This release intentionally keeps localStorage as a fallback so the dashboard does not become unusable before Supabase is configured.

The SQL currently allows anonymous access to the single dashboard row only as a temporary setup step. **Before putting the dashboard into real multi-user production use, add Login + Admin / Team Lead / Member roles and replace the temporary policies with authenticated role-based policies.**
