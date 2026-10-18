# Hostinger deployment notes (V51)

This ZIP adds a `build` script to `server/package.json` so `npm run build` no longer fails with `Missing script: build`.

## Important: choose the correct Hostinger deployment type

The screenshot provided showed **Framework preset: Vite** while the root directory was `server`. That is the wrong framework preset for an Express API. The added build script resolves the missing-script error, but it does **not** turn the Express backend into a Vite/static website.

For the backend, use a Hostinger Node.js application/runtime deployment that supports a start command. Set:

- Root directory: `server`
- Framework/runtime: Node.js (not Vite), if available
- Package manager: `npm`
- Install command: `npm install`
- Build command: `npm run build`
- Start command: `npm start`
- Node version: `22.x` if supported by the plan

If Hostinger's form requires an output directory such as `dist`, the build script now creates a small `server/dist` marker folder. **This output folder is not the Express app itself**. Only proceed if the selected Hostinger Node.js runtime launches the app using `npm start`. If the deployment is a Vite/static-only deployment and does not support a Node.js start command, it cannot host this backend with these settings; use Hostinger's Node.js Web App/VPS option or keep the backend on Render.

## Frontend deployment (separate app)

Create a second deployment from the same repository:

- Root directory: `client`
- Framework: Vite
- Install command: `npm install`
- Build command: `npm run build`
- Output directory: `dist`
- Build-time environment variable: `VITE_API_URL=https://YOUR-BACKEND-HOSTNAME`

Replace the example URL with the actual backend HTTPS URL. The backend CORS variable `CLIENT_ORIGIN` must match the frontend origin exactly, for example `https://trokotakavach.com` (use your actual final domain).

## Backend environment variables

Configure these in Hostinger's environment-variable panel, using the values from your own secure records:

- `NODE_ENV=production`
- `MONGODB_URI`
- `JWT_SECRET`
- `UPLOAD_PASSWORD`
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`
- `CLIENT_ORIGIN`
- SMTP variables only if email/password-reset functionality is configured

Never commit secrets to GitHub or include them in this ZIP. Since a prior screenshot exposed secret values, rotate the admin password, JWT secret, and MongoDB database-user password before going live.

## Smoke tests

After deployment, check runtime logs and verify the server is listening on `process.env.PORT` (the existing server code uses `process.env.PORT || 5000`). Test login, database reads/writes, Excel dropdowns/uploads, exports, and CORS from the deployed frontend.

## Source note

This ZIP preserves the existing V50 application source and current-date fix; it only adds the backend build script and this deployment guide. A successful `npm run build` alone does not verify that Hostinger's selected hosting product supports the Express runtime.
