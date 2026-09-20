# Traction Operation - Kavach Monitoring System

MERN/Vite application for West Central Railway, Kota Division.

## Architecture
- `client`: React + Vite
- `server`: Node.js + Express
- MongoDB Atlas
- Excel source workbooks: `TRAINS_PER_DAY.xlsx` and `Kavach_Loco_Details.xlsx`

## Authentication
The application now uses an HttpOnly JWT cookie for authenticated users. The server checks the user and `authVersion` on every protected request. Logout increments `authVersion`, clears the cookie, and therefore invalidates the current authentication state server-side.

Roles:
- `admin`: Dashboard, Stored Data / Export, Summary, Data Sources, Password Reset Requests
- `shed`: Shed Remark
- `oem`: OEM Remark

Shed and OEM accounts are synchronized from `Kavach_Loco_Details.xlsx` (`SHED` and `Loco_Kavach_Make`). New accounts start with the temporary password `123`.

## Environment variables

Copy `server/.env.example` to the deployment environment and set real values. Never commit real secrets.

Required:
- `MONGODB_URI`
- `JWT_SECRET`
- `UPLOAD_PASSWORD`
- `CLIENT_ORIGIN`
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`

Optional SMS:
- `TWILIO_ACCOUNT_SID`
- `TWILIO_AUTH_TOKEN`
- `TWILIO_FROM_NUMBER`

Frontend:
- `VITE_API_URL`

## Local setup

### Server
```bash
cd server
npm install
npm start
```

### Client
```bash
cd client
npm install
npm run dev
```

Open the Vite URL, normally `http://localhost:5173`.

## Production notes
- Set `NODE_ENV=production`.
- Set `CLIENT_ORIGIN` to the exact frontend origin(s), comma separated if required.
- Use HTTPS for both frontend and backend.
- Keep MongoDB credentials, JWT secret, admin credentials, upload password, and SMS credentials out of source control.
- The frontend uses `withCredentials: true` for the authentication cookie.
- If the SMS provider is not configured, an approved reset still changes the password and the Admin UI shows that SMS was not sent/configured.


## Production CORS / Cookie requirement
The frontend uses `withCredentials: true` because authentication is stored in an HttpOnly cookie. The Express API therefore enables credentialed CORS and allows the deployed Vercel origin explicitly. Set `CLIENT_ORIGIN` on Render to the deployed frontend origin (for example `https://traction-operation.vercel.app`). Do not use `*` for a credentialed CORS origin.

## Latest fixes
- Summary count links now navigate inside the Vercel SPA without a full-page reload, so clicking a count no longer lands on a Vercel 404 page.
- Vercel SPA rewrites are included in `client/vercel.json`, so `/dashboard`, `/stored-data`, `/summary`, `/shed-remark`, `/oem-remark`, and other client routes also survive a browser refresh/direct URL.
- Summary count links open the exact clicked date + DIR + metric; the Back to Summary action restores the original Summary date range + DIR.
- Browser Back and the in-page `← Back to Summary` flow are preserved without forcing a page reload.
- Shed Remark and OEM Remark now show the complete dashboard dataset, including both `Shed_Remark` and `OEM` columns, with the existing role-specific filters.
- Shed/OEM remark rows now start in Update mode when no remark exists; after the first successful save the action changes to Modify. Modify re-enables only that role's remark field and saves the same MongoDB record without creating a duplicate.
- Responsive behavior remains horizontal-scroll-inside-the-table on narrow screens; the overall page does not gain horizontal scrolling.


## Latest fixes
- Login now re-runs the user bootstrap so newly uploaded Shed/OEM Excel users are provisioned without a server restart. Shed users are collected from both TRAINS_PER_DAY.xlsx and Kavach_Loco_Details.xlsx; OEM users are collected from Kavach_Loco_Details.xlsx.
- Temporary/reset accounts (`mustChangePassword`) are kept on the temporary password until the user completes Change Password; accounts that already changed their password are not overwritten.
- Shed Remark page shows every Dashboard column except `OEM`, plus ACTION; `Shed_Remark` is the editable role field.
- OEM Remark page shows every Dashboard column except `Shed_Remark`, plus ACTION; `OEM` is the editable role field.
- Both role pages retain Update -> Modify -> Update behavior for the role-specific remark.


## V6 UI refinement
- Shed Remark and OEM Remark tables intentionally omit `Day`, `Loco_attached_division`, `Loco Link`, `Loco link division`, `Trn type`, `loco type`, and `Rly` to reduce horizontal width and preserve responsiveness.
- The displayed date column on both remark pages is labeled `Date run Kota div`.
- `SHED_Remark` is positioned immediately after the date on the Shed Remark page.
- Remark tables retain the Dashboard-style internal scrolling container without forcing a 3000px-wide table.
- Update/Modify controls are compacted across applicable operational pages; `.close-btn` is styled consistently wherever that control is used.
