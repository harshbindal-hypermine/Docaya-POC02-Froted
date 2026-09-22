# DocayaGithubPOC02

# Docaya | دوكايا

**Document Management Intelligence · ذكاء إدارة الوثائق**

An interactive English/Arabic document management prototype for UAE Ministry of Interior client workshops. Explore document capture, approvals, publishing, discovery, records governance and administration with synthetic MOI scenarios.

**[Open the live demo](https://sthmaihrnhp04.z13.web.core.windows.net/docayapro/)** · [Client demonstration guide](docs/prototype-demo.md) · [Architecture](docs/architecture.md)

<a href="public/brand/docaya-welcome.png"><img src="public/brand/docaya-welcome.png" width="680" alt="Selected Docaya identity: a three-dimensional emerald D folder with white sheets and a pale UAE skyline"></a>

Select **Enter demo workspace** and use **System Administrator** to explore all modules. Switch language to demonstrate Arabic and right-to-left layouts.

> This is a browser-based UI/UX prototype. Sample records and identities are fictional; AI, authentication and external integrations are simulated. No backend, Azure account, credentials or `.env` file is required to run it.

## Approved branding

The application uses the following selected images consistently across the login screen and workspace:

| Identity          | Selected image                                                                     | Application placement                                                                             |
| ----------------- | ---------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| **Docaya**        | [Emerald D folder with documents and UAE skyline](public/brand/docaya-welcome.png) | Full artwork on login, a compact 60 px view in the top-left corner, favicon and Apple touch icon. |
| **Ask Docaya AI** | [D speech bubble with white sparkle](public/brand/ask-docaya-mark.png)             | Assistant launcher, panel, responses, search and discovery.                                       |

See the [login screenshot](docs/images/docaya-login.png), [brand guide](docs/branding.md) and [latest release notes](docs/release-notes.md). Earlier logo explorations remain in the repository as reference assets.

## Latest experience

| Area                       | What you can demonstrate                                                                                                       |
| -------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| Document registration      | Upload an original or choose a sample, review its preview and complete Capture, Classify and Review & submit.                  |
| Automatic classification   | Inspect local suggestions and their rationale, override fields and explicitly confirm classification in the registration form. |
| Approval queue             | Scan pending documents by due date, open a focused review drawer, approve or return with a reason, and continue explicitly.    |
| Search & Ask Docaya        | Browse document thumbnails or a list, filter results, save queries and open cited records.                                     |
| Docaya AI Assistant        | Open the floating bottom-right assistant, inspect source thumbnails and provide response feedback.                             |
| UAE design and readability | Selected 3D emerald D Docaya identity and D speech-bubble AI logo, UAE-inspired colours, bundled Inter and Noto Sans Arabic.   |
| Governance and workshops   | Explore holds, retention, correspondence, reports, audit history and M1–M15 coverage; record and export client expectations.   |

![Docaya document register with readable typography and synthetic UAE MOI records](docs/images/docaya-register.png)

![Docaya approval inbox with pending documents, departments and review due dates](docs/images/docaya-approvals.png)

![Docaya approval drawer with the queue, document preview and review details visible together](docs/images/docaya-review.png)

**Approvals** opens a full-width **To review** list containing pending documents only, with the earliest due first. Search documents or select **Overdue only**; **Returned** is a separate view for documents awaiting owner changes. Registration remains available from the document register.

Opening a document changes the desktop list to a compact queue beside its preview and **Review**, **Details** and **Activity** tabs. One **Approve & sign** action and one **Return for changes** action keep the decision clear. Returning requires an explanation and **Confirm return**; **Cancel** leaves the document unchanged. Feedback explains the saved decision, and **Next pending document** stays within the filtered queue. Partial approvals require **Review the remaining step** before another decision. The app does not automatically approve or advance.

The review drawer has no blurred backdrop. On narrow screens the content stacks; closing the drawer returns to the queue. The assistant is hidden while an approval drawer is open so it cannot cover the decision controls. Other document-management screens retain their existing overview, workflow, governance and activity tabs.

## Run locally

Prerequisites: **Node.js 22.13–26** and **npm 10–11**. Run these commands from the repository root:

```powershell
npm ci
npm start -- --strictPort
```

Open **[http://localhost:5173/](http://localhost:5173/)** and select **Enter demo workspace**. Keep the terminal running; press `Ctrl+C` to stop the server.

If port 5173 is already occupied, use another port:

```powershell
npm start -- --port 5175 --strictPort
```

Then open **http://localhost:5175/**. Use `Ctrl+F5` if the browser shows older assets after an update. The language toggle changes English/Arabic; `Ctrl+J` on Windows or `Cmd+J` on macOS toggles the assistant when it is available in the workspace, outside an approval or registration drawer.

## SharePoint Online Integration

Docaya includes a backend-only Microsoft Graph integration for SharePoint Online. The browser calls Docaya API routes under `/api/v1/sharepoint/*`; it never receives Microsoft Graph tokens, client secrets or direct Graph endpoints.

Configure the backend with an Entra app registration that has Microsoft Graph **Application** permissions for `Sites.Read.All` and `Files.ReadWrite.All`, then grant admin consent in Entra ID. Create a client secret for local/server use and keep it in `.env` or the deployment secret store only.

Required local placeholders:

```powershell
TENANT_ID=<Entra tenant ID>
CLIENT_ID=<Entra application client ID>
CLIENT_SECRET=<client secret value>
SHAREPOINT_HOSTNAME=hypermineai.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/DocayaDocLibrary-Dev
SHAREPOINT_DRIVE_NAME=Shared Documents
```

The target SharePoint site is `https://hypermineai.sharepoint.com/sites/DocayaDocLibrary-Dev` and the document library is `Shared Documents`. Legacy `SHAREPOINT_SITE_URL` and `SHAREPOINT_DRIVE_ID` are still accepted for managed-identity deployments, but new configuration should prefer hostname, site path and drive name so the backend proves it resolved the configured site and library at startup/readiness time.

Do not commit `.env`, client secrets, certificates or exported app credentials. A working connection is only established after the backend successfully authenticates with Microsoft Graph and resolves both the configured SharePoint site and the `Shared Documents` drive.

## Client demo route

1. **Overview and document register:** introduce Docaya, synthetic MOI records and Arabic layout.
2. **Register document:** choose a traffic, civil defence or evidence sample; inspect classification, override a field and confirm your review.
3. **Approvals → To review:** scan due dates, open a document, approve or confirm a return with a reason, then choose the next pending document. Use **Returned** to inspect owner feedback.
4. **Staging & publishing → Search indexing → Run demo batch:** index published records, then demonstrate thumbnail search and the floating assistant.
5. **Records & retention, Correspondence and Client workshop:** demonstrate governance scenarios and capture the client's expectations; export workshop notes to CSV.

The [detailed demonstration guide](docs/prototype-demo.md) includes a fifteen-minute walkthrough, preview behavior and the implemented/future scope for all fifteen modules.

## Data and prototype boundaries

- **Local workspace:** documents, settings and workshop notes persist in localStorage; uploaded originals are stored in IndexedDB. Localhost and the hosted site have separate workspaces. Git pushes and deployments do not synchronize browser data.
- **Uploads:** PDF, DOCX, XLSX, PPTX, PNG, JPEG, TIFF and TXT, up to 20 MB per file. Hashing, basic file-header checks and duplicate detection are demonstrated; these are not malware or DLP scanning.
- **Classification:** local rules inspect filenames and up to the first 12,000 characters of TXT files. Suggestions are editable, and confidence indicates a rule score. There is no PDF/Office text extraction, image OCR or model inference.
- **Discovery:** retrieval matches registered metadata and summaries from accessible, published and indexed records. Answers cite those records. It does not search the full contents of arbitrary uploads or call a live LLM.
- **Previews:** supported images, TXT and browser-supported PDF originals can be previewed. Office files and PDF thumbnails use covers. Seeded documents use marked sample previews and downloadable sample text.
- **Integrations:** identity, signatures, publishing destinations, notifications and correspondence demonstrate local state transitions. Browser role controls are not production authorization, and no external message is sent.
- **Reset:** export workshop notes before resetting demo data. Reset removes this prototype's local records, uploaded originals and notes and restores the sample workspace.

## Validate changes

```powershell
npm run build
npm run lint
npm test
npm run test:e2e
```

The prototype browser journeys cover classification, pending queue ordering and filters, approval and return decisions, search, governance, correspondence, administration and Arabic accessibility. The local configuration uses installed Chrome; CI installs Chromium. See [playwright.prototype.config.ts](playwright.prototype.config.ts).

To run those browser journeys against the hosted build:

```powershell
npx cross-env DOCAYA_E2E_URL=https://sthmaihrnhp04.z13.web.core.windows.net/docayapro/ npm run test:e2e
```

The latest UI revision was also checked for loaded bilingual fonts, desktop/mobile layout, review and registration drawers, and automated accessibility findings. These checks support prototype quality; they do not establish production compliance.

## Publish the hosted demo (Azure Storage static website)

Canonical repository: **[hypermine2050/DocayaGithubPOC02](https://github.com/hypermine2050/DocayaGithubPOC02)**.

Git remote `origin` should point to:

```powershell
https://github.com/hypermine2050/DocayaGithubPOC02.git
```

The hosted demo URL is **[https://sthmaihrnhp04.z13.web.core.windows.net/docayapro/](https://sthmaihrnhp04.z13.web.core.windows.net/docayapro/)**, served from the `$web/docayapro/` folder of the Azure Storage account `sthmaihrnhp04` (resource group `rg-hmai-hypersigma-hp03`, subscription `hmai-hyperaihub-products`). The storage account also hosts other applications at the site root, so Docaya must only ever be published inside its own `docayapro/` folder.

To deploy an update:

```powershell
# 1. Build with the hosted base path
$env:DOCAYA_BASE_PATH = '/docayapro/'
npm run build

# 2. Upload dist/ into $web/docayapro/ with AzCopy and a short-lived SAS token
#    (generate the SAS in the Azure Portal: Blob service, Container + Object
#    resource types, Read/Write/Add/Create/List permissions)
azcopy copy "dist\*" 'https://sthmaihrnhp04.blob.core.windows.net/$web/docayapro/?<SAS>' --recursive
```

Do not overwrite files outside `docayapro/`. Paste SAS tokens directly into the terminal and let them expire quickly.

> **GitHub Pages is not used.** The repository is private on a plan without Pages support, so the [Publish prototype to GitHub Pages](.github/workflows/prototype-pages.yml) workflow cannot deploy. [Prototype quality](.github/workflows/ci.yml) still validates every push to `main`.

For a local production-build preview:

```powershell
npm run build
npm run preview -- --host localhost --port 4173 --strictPort
```

Open **http://localhost:4173/**. Each origin has independent demo data, including the production preview port.

## Documentation

| Document                                                                    | Purpose                                                                                               |
| --------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| [Client demonstration guide](docs/prototype-demo.md)                        | Walkthrough, current functionality, typography, persistence and M1–M15 coverage.                      |
| [Release notes](docs/release-notes.md)                                      | Approved branding, current demo improvements and validation of the published application.             |
| [Brand assets and usage](docs/branding.md)                                  | Docaya and AI logo gallery, welcome artwork, palette, file links and presentation guidance.           |
| [Architecture](docs/architecture.md)                                        | Current browser implementation, source map, storage and publishing; retained connected-system design. |
| [Enterprise specification](<docs/Solution Doc/Docaya-DMS-Specification.md>) | Original enterprise requirements and engineering baseline, with a current prototype status note.      |
| [Threat model](docs/threat-model.md)                                        | Retained connected-application security design reference.                                             |
| [Operations runbook](docs/runbooks/operations.md)                           | Retained API and infrastructure operations reference.                                                 |
| [API contract](docs/api/openapi.json)                                       | Retained connected-application API contract.                                                          |

The default entry point is `src/main.tsx` → `src/prototype/PrototypeApp.tsx`. Earlier `src/App.tsx`, `server/` and infrastructure files remain in the repository as connected-application references; the prototype startup commands do not activate that application or its organizational sign-in.
