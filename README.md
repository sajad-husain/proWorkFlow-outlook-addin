# proWorkFlow Outlook Add-in

Outlook add-in that lets you create tasks in ProWorkflow directly from Outlook. Extracts email context (subject, body, sender) and pre-fills a task creation form — with mock mode and real API support.

## Project Structure

```
proWorkFlow/
├── src/
│   ├── commands/                        # Ribbon command functions (hidden page)
│   │   ├── commands.html
│   │   └── commands.ts
│   └── taskpane/                        # Main task pane UI (React SPA)
│       ├── index.tsx                    # React entry point
│       ├── taskpane.html                # HTML shell
│       ├── taskpane.ts                  # Office.js utilities
│       ├── hooks/
│       │   └── useEmailContext.ts       # Extract email data from Outlook item
│       ├── services/
│       │   ├── proworkflowApi.ts        # Axios-based ProWorkflow API client (mock + real)
│       │   └── mockData.ts             # Mock data for API responses
│       └── components/
│           ├── App.tsx                  # Root component
│           ├── Header.tsx               # Logo & title
│           ├── CreateTask/
│           │   └── CreateTaskForm.tsx   # Full task creation form
│           └── Layout/
│               └── AppLayout.tsx        # MUI drawer layout (for future routing)
├── manifest.xml                         # Outlook add-in registration
├── webpack.config.js                    # Build config (multi-entry)
├── tsconfig.json                        # TypeScript config
├── package.json                         # Dependencies & scripts
├── assets/                              # Icons & images
└── docs/                                # Documentation (see below)
```

## Documentation

Detailed documentation for developers and LLMs:

| Document | Audience | Contents |
|---|---|---|
| [docs/ARCHITECTURE.md](proWorkFlow/docs/ARCHITECTURE.md) | All | System architecture, component tree, tech stack, boot sequence, manifest structure |
| [docs/CODE_FLOW.md](proWorkFlow/docs/CODE_FLOW.md) | All | Complete execution paths, data flow diagrams, import/export graph, error handling |
| [docs/DEBUGGING.md](proWorkFlow/docs/DEBUGGING.md) | Developers | Debug setup, VS Code configs, troubleshooting, available scripts & tasks |
| [docs/FILE_MAP.md](proWorkFlow/docs/FILE_MAP.md) | LLMs | File-by-file reference with exports, imports, key lines, and dependency graph |

## Quick Start

```powershell
cd proWorkFlow
npm install
npm start            # Build + sideload + launch Outlook
```

Press **F5** in VS Code (with "Outlook Desktop (Edge Chromium)" launch config) to start debugging.

## Tech Stack

- **React 18.2** + **TypeScript 5.4** (ES5 target)
- **Fluent UI React v9** (theming, components, CSS-in-JS)
- **MUI Material v9** (layout components — AppLayout drawer)
- **Axios** (ProWorkflow API client, mock mode by default)
- **React Router DOM v7** (routing for future multi-view)
- **Office.js** (mailbox integration via CDN)
- **Webpack 5** (multi-entry bundling, HMR, HTTPS dev server)
- **core-js** + **regenerator-runtime** (IE11 polyfills)

## Development Scripts

| Script | Purpose |
|---|---|
| `npm start` | Build + sideload + launch Outlook |
| `npm stop` | Stop debugging + remove sideload |
| `npm run dev-server` | Start dev server only (localhost:3000) |
| `npm run build` | Production build |
| `npm run build:dev` | Development build (source maps) |
| `npm run watch` | Rebuild on file changes |
| `npm run lint` | Run ESLint |
| `npm run validate` | Validate manifest.xml |
| `npm run signin` | M365 account login |
| `npm run signout` | M365 account logout |
| `npm run lint:fix` | Auto-fix lint issues |
| `npm run prettier` | Format with Prettier |
