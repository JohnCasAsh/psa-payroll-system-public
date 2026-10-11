# PSA Payroll System

A payroll and HR system built for the Philippine Statistics Authority, Regional Statistical Services Office II (RSSO II).

This public copy contains the source code only. It holds **no real employee data**; the names used as examples (for instance "Maria L. Santos") are made up.

## Features

- **Employees:** directory, compensation, deductions and loans, and position history.
- **Pay slips and Payroll Archive:** quincena payslips, a payroll register in Excel, saved payroll snapshots, and an annual breakdown.
- **Leave (CS Form No. 6):**
  - The employee files; HR certifies leave credits (7.A); the employee's division Chief recommends (7.B); the Regional Director approves or disapproves.
  - Each officer sees an application only at their own step.
  - Signatures are drawn on screen.
  - Pop-up notifications go to the next officer and to the applicant.
- **201 Files:** employees keep their own document records.
- **Roles and permissions:** a role manager with fine-grained permissions, and a superadmin account.
- **Settings:** payslip signatories, leave signatories (with division routing), deduction and income types, salary grades and divisions.

## Tech stack

- **Frontend:** React and Vite, in `apps/frontend`.
- **Backend:** Node.js and Express, in `services/api`; the entry point is `src/server.mysql.js`.
- **Database:** MySQL or MariaDB. The tables are created automatically when the API starts.

## Run it locally (Windows + XAMPP)

1. **Database:** start MySQL in XAMPP and create an empty database named `psa_payroll`.

2. **Backend settings:** create `services/api/.env`:

   ```env
   JWT_SECRET=replace-with-a-long-random-string
   ADMIN_USERNAME=admin
   DB_HOST=127.0.0.1
   DB_PORT=3306
   DB_NAME=psa_payroll
   CORS_ORIGIN=http://localhost:5173
   ```

   This setup connects as MySQL `root` with no password (the XAMPP default). To use another account, set `MYSQL_URL=mysql://user:password@host:3306/psa_payroll` instead.

3. **Start the backend:**

   ```bash
   cd services/api
   npm install
   node --env-file=.env src/server.mysql.js
   ```

4. **Create the superadmin:** run this in a second terminal. It uses `ADMIN_USERNAME` from `.env`.

   ```bash
   cd services/api
   node --env-file=.env seed-admin.mjs "YourStrongPassword!1"
   ```

5. **Start the frontend:**

   ```bash
   cd apps/frontend
   npm install
   npx vite --port 5173 --host 0.0.0.0
   ```

   Open <http://localhost:5173/login.html> and sign in as the superadmin.

6. **Set up the office:** in **Settings**, create roles and give them permissions, for example HR (*HR certification*), Chief (*Chief recommendation*) and Director (*Director approval*). Then link each user to their employee record, and set the leave signatories with their divisions under **Settings → Leave Signatories**.

New users register on the sign-in page. An administrator approves them in **Settings → User Management**.

## Docker (optional)

`docker-compose.mysql.yml` runs MySQL, Redis, MinIO, the API and the frontend. It reads `JWT_SECRET` and `MINIO_SECRET_KEY` from a `.env` file next to it, so set both first:

```bash
docker compose -f docker-compose.mysql.yml up --build
```

## Notes

- Never commit a real `.env` or exported payroll files. `.gitignore` already excludes `.env*` and `*.sql`.
- This system handles personal and financial data. Follow the Data Privacy Act of 2012 (RA 10173) when you deploy it with real records.
