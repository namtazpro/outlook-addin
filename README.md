# outlook-addin

Prototype of an **Outlook add-in** integrated with an **Azure Functions** back end for the
*Contoso Consents Management (CCM)* scenario. While reading an email, a user picks one of
their CCM projects and uploads the email (body + attachments) to Azure Blob Storage via an
HTTP-triggered function.

The repository contains two distinct kinds of content:

- **Working implementation** — the code that exists and runs today, under
  [ContosoConsentsManagement/](ContosoConsentsManagement/) (add-in) and
  [func_contosoconsentsmgmt/](func_contosoconsentsmgmt/) (Functions back end).
- **Post-hoc research** — design notes under [research/](research/) produced **after** the
  prototype was built. These documents explore the broader target architecture (multi-tenant
  Entra ID SSO, cross-client coverage, distribution, etc.) and describe what a productized
  solution should look like; they do **not** describe the current implementation.

## Repository layout

```text
outlook-addin/
├── ContosoConsentsManagement/   # Outlook add-in (React + TypeScript, Office.js)
├── func_contosoconsentsmgmt/    # Azure Functions back end (.NET 8, isolated worker)
└── research/                    # Architecture & design research notes
```

### `ContosoConsentsManagement/` — Outlook add-in

React 18 + Fluent UI v9 task pane add-in scaffolded from
`OfficeDev/Office-Addin-TaskPane-React`, surfaced on the **message read** ribbon in Outlook.

Key files:

- [ContosoConsentsManagement/manifest.xml](ContosoConsentsManagement/manifest.xml) — Mail add-in manifest (XML), declares the `MessageReadCommandSurface` button and task pane.
- [ContosoConsentsManagement/package.json](ContosoConsentsManagement/package.json) — Build, lint, and `office-addin-debugging` scripts (`npm start`, `npm run dev-server`, `npm run validate`).
- [ContosoConsentsManagement/webpack.config.js](ContosoConsentsManagement/webpack.config.js) — Dev server on `https://localhost:3000` with HTTPS certs from `office-addin-dev-certs`.
- `src/taskpane/`
  - [index.tsx](ContosoConsentsManagement/src/taskpane/index.tsx) / [taskpane.html](ContosoConsentsManagement/src/taskpane/taskpane.html) — Task pane entry point.
  - [config.ts](ContosoConsentsManagement/src/taskpane/config.ts) — Endpoint URLs (`UPLOAD_ENDPOINT_URL`, `PROJECTS_ENDPOINT_URL`) and dev mock projects.
  - [emailPayload.ts](ContosoConsentsManagement/src/taskpane/emailPayload.ts) — Builds the JSON payload (headers, body, attachments as base64) from `Office.context.mailbox.item`.
  - [projectsApi.ts](ContosoConsentsManagement/src/taskpane/projectsApi.ts) — Fetches the user's allowed project list (falls back to `MOCK_PROJECTS` in dev).
  - `components/` — `App.tsx`, `Header.tsx`, and the main [SendEmailToBlob.tsx](ContosoConsentsManagement/src/taskpane/components/SendEmailToBlob.tsx) view (project dropdown + upload button + status).
- `src/commands/` — [commands.ts](ContosoConsentsManagement/src/commands/commands.ts) / [commands.html](ContosoConsentsManagement/src/commands/commands.html) for ribbon command function.
- `assets/` — Add-in icons (16/32/64/80/128) and logos.

> Note: `config.ts` currently embeds a function key in the upload URL. Before sharing or
> deploying publicly, move that secret out of source (env-based config / build-time
> injection) and rotate the key.

### `func_contosoconsentsmgmt/` — Azure Functions back end

.NET 8 isolated-worker Functions app receiving the add-in's payload.

- [InsertContent.cs](func_contosoconsentsmgmt/InsertContent.cs) — `POST /api/InsertContent` (function-key auth). Parses the JSON envelope, builds a deterministic blob name from `internetMessageId` + `capturedAt`, and writes the email JSON plus extracted attachments to a Blob container using a SAS URL from the `BlobContainerSasUrl` app setting. Correlates envelope and attachments via a generated correlation id.
- [Program.cs](func_contosoconsentsmgmt/Program.cs) — Functions host bootstrap with Application Insights worker telemetry.
- [host.json](func_contosoconsentsmgmt/host.json), [local.settings.json](func_contosoconsentsmgmt/local.settings.json), [func_contosoconsentsmgmt.csproj](func_contosoconsentsmgmt/func_contosoconsentsmgmt.csproj) — Host configuration and project file.
- [Properties/launchSettings.json](func_contosoconsentsmgmt/Properties/launchSettings.json) — Local F5 profile.
- `tests/`
  - [InsertContent.http](func_contosoconsentsmgmt/tests/InsertContent.http) — REST Client requests for local and deployed runs.
  - [Invoke-InsertContent.ps1](func_contosoconsentsmgmt/tests/Invoke-InsertContent.ps1) — PowerShell test harness.
  - [sample-payload.json](func_contosoconsentsmgmt/tests/sample-payload.json) — Example email payload.
- `bin/`, `obj/` — Build output (normally gitignored).

### `research/` — Design notes (produced after the prototype)

Forward-looking design research written **after** the code in
`ContosoConsentsManagement/` and `func_contosoconsentsmgmt/` was implemented. It captures
what an end-to-end, multi-tenant productized version of this integration should look like
and is **not** a description of the current prototype. Expect divergence between these
documents and the code (for example, the prototype uses a function-key-protected HTTP
endpoint, while the research recommends Entra ID SSO via Nested App Authentication).

- [research/2026-04-29/outlook-addin-ccm-integration-research.md](research/2026-04-29/outlook-addin-ccm-integration-research.md) — Consolidated task research: surface/runtime, manifest strategy, cross-client matrix, multi-tenant Entra ID auth (NAA + Office Dialog/MSAL.js fallback), Office.js + Graph capture flow, CCM API contract, distribution, non-Outlook clients.
- `research/subagents/2026-04-29/`
  - [outlook-addin-platform-research.md](research/subagents/2026-04-29/outlook-addin-platform-research.md) — Add-in platform & runtime details.
  - [auth-multitenant-sso-research.md](research/subagents/2026-04-29/auth-multitenant-sso-research.md) — Multi-tenant SSO options.
  - [deployment-and-other-clients-research.md](research/subagents/2026-04-29/deployment-and-other-clients-research.md) — AppSource / Integrated Apps / Gmail coverage.

## End-to-end flow

1. User opens an email in Outlook and clicks the **Contoso Consents Management** ribbon button.
2. Task pane loads the user's projects via `projectsApi.fetchProjects` (or `MOCK_PROJECTS` in dev).
3. User selects a project and clicks **Upload**; `emailPayload.buildEmailPayload` collects subject, body, headers, and base64-encoded attachments.
4. The payload is `POST`ed to the Function's `UPLOAD_ENDPOINT_URL`.
5. `InsertContent` validates the JSON, derives a deterministic blob name, and writes the envelope + attachments to Blob Storage using `BlobContainerSasUrl`.

### What happens when the add-in starts

When the user clicks the ribbon button (or selects an email while the pane is pinned),
Outlook activates the add-in based on the rules declared in
[manifest.xml](ContosoConsentsManagement/manifest.xml) (`MessageReadCommandSurface`,
activation rule `ItemIs ItemType="Message"`):

1. **Host loads the task pane.** Outlook opens the URL declared in the manifest's
   `SourceLocation` (in dev: `https://localhost:3000/taskpane.html`, served by webpack-dev-server
   with the certificates produced by `office-addin-dev-certs`). The page bootstraps
   `Office.js`, then [index.tsx](ContosoConsentsManagement/src/taskpane/index.tsx) waits for
   `Office.onReady` before rendering the React tree
   ([App.tsx](ContosoConsentsManagement/src/taskpane/components/App.tsx) →
   [Header.tsx](ContosoConsentsManagement/src/taskpane/components/Header.tsx) +
   [SendEmailToBlob.tsx](ContosoConsentsManagement/src/taskpane/components/SendEmailToBlob.tsx)).
2. **Identity is read locally.** `SendEmailToBlob` reads
   `Office.context.mailbox.userProfile.emailAddress` synchronously to identify the signed-in
   Outlook user. No interactive sign-in or token acquisition happens at this stage — the
   prototype currently relies on the Functions function key for backend auth (this is one of
   the gaps called out in `research/`).
3. **Projects are fetched.** A `useEffect` calls
   [projectsApi.fetchProjects](ContosoConsentsManagement/src/taskpane/projectsApi.ts) with that
   email. If `PROJECTS_ENDPOINT_URL` is set, it issues a `GET` against the projects API;
   otherwise it returns `MOCK_PROJECTS` from
   [config.ts](ContosoConsentsManagement/src/taskpane/config.ts). While the request is in
   flight the dropdown shows a `Spinner`; on failure a `MessageBar` surfaces the error and the
   upload button stays disabled.
4. **Pane becomes interactive.** Once projects resolve, the dropdown is populated and the
   **Send email to Blob Storage** button enables as soon as a project is selected. The
   currently-selected message in Outlook is the implicit target — `buildEmailPayload` reads
   it later from `Office.context.mailbox.item` only when the user clicks Upload, so switching
   emails simply changes which message will be captured on the next click.
5. **Ribbon command surface.** The button itself runs through
   [commands.html](ContosoConsentsManagement/src/commands/commands.html) /
   [commands.ts](ContosoConsentsManagement/src/commands/commands.ts), which is the
   `FunctionFile` declared in the manifest. For this prototype it just opens the task pane;
   no UI-less ribbon action is performed.

## Getting started

Add-in (from `ContosoConsentsManagement/`):

```powershell
npm install
npm run validate   # validate manifest.xml
npm start          # sideload into Outlook desktop
```

Functions back end (from `func_contosoconsentsmgmt/`):

```powershell
# Set BlobContainerSasUrl in local.settings.json
func start
# or F5 in VS Code, then use tests/InsertContent.http / Invoke-InsertContent.ps1
```

## Configuration

No secrets are committed to source. Both projects need local configuration before
they will run.

### Add-in (`ContosoConsentsManagement/`)

Endpoints are injected at build time by webpack from environment variables (or a local
`.env` file). See [`ContosoConsentsManagement/.env.example`](ContosoConsentsManagement/.env.example).

1. Copy the template:

   ```powershell
   cd ContosoConsentsManagement
   Copy-Item .env.example .env
   ```

2. Edit `.env` and set:

   | Variable | Required | Description |
   | --- | --- | --- |
   | `UPLOAD_ENDPOINT_URL` | yes | Full URL of the `InsertContent` Function, **including** `?code=<function key>`, e.g. `https://<your-func-app>.azurewebsites.net/api/InsertContent?code=...`. |
   | `PROJECTS_ENDPOINT_URL` | no | Projects API base URL. Leave empty to fall back to `MOCK_PROJECTS` in [config.ts](ContosoConsentsManagement/src/taskpane/config.ts). |

3. Rebuild / restart the dev server (`npm start`) so the values are baked in.

`.env` is gitignored. Real shell environment variables (e.g. set in CI) take
precedence over the `.env` file.

### Functions back end (`func_contosoconsentsmgmt/`)

Settings are read from `local.settings.json` locally and from App Settings in Azure.
`local.settings.json` is gitignored.

Create `func_contosoconsentsmgmt/local.settings.json` with at least:

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
    "BlobContainerSasUrl": "https://<storage>.blob.core.windows.net/<container>?<sas-token>"
  }
}
```

| Setting | Required | Description |
| --- | --- | --- |
| `BlobContainerSasUrl` | yes | Full SAS URL of the destination Blob container (write permission). Used by [InsertContent.cs](func_contosoconsentsmgmt/InsertContent.cs). |
| `AzureWebJobsStorage` | yes | Functions runtime storage. `UseDevelopmentStorage=true` works locally with Azurite. |
| `FUNCTIONS_WORKER_RUNTIME` | yes | Must be `dotnet-isolated`. |

In Azure, set the same values under **Function App → Configuration → App settings**.

## Deployment

Deploy the Azure Function to your own subscription, resource group, and Function App.
Replace the placeholders in `.env` (add-in) and the Function App's application settings
(`BlobContainerSasUrl`) with values for your environment.

