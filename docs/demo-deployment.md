# Demo deployment

The repository includes a Node.js demo server and a synthetic SVG welcome screen. No public demo deployment is configured or claimed. The demo reads a selected file in the browser; it does not upload that file to the server.

## Run locally

Use Node.js 22, matching CI:

```sh
npm ci
npm run build
npm run demo
```

Open <http://127.0.0.1:4173/>. By default, the server listens only on loopback at port `4173`, as before.

## Deploy to a Node.js host

Use a Node.js host that can build from the repository and run a long-lived Node process. Configure the service with:

- **Repository:** `dukeofwest/universal-file-preview`
- **Runtime:** Node.js 22
- **Build command:** `npm ci && npm run build`
- **Start command:** `npm run demo`
- **Host setting:** `HOST=0.0.0.0`
- **Port setting:** use the port assigned by the hosting provider as `PORT`; do not hard-code a public port if the provider supplies one.

The server validates `PORT` and keeps the existing `127.0.0.1:4173` default for local development. After the provider reports a deployment URL, verify the root page, select representative synthetic or properly licensed sample files, and check the browser console and network panel. Do not upload private user documents as demo fixtures. Add a public demo link to the README and npm release only after the deployment URL and its behavior have been verified.

Deployment requires an account and configuration in the hosting provider. This repository does not contain provider credentials or an auto-deploy workflow.
