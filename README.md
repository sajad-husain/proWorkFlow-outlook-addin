# proWorkFlow Outlook Add-in

Outlook add-in that lets you create tasks in ProWorkflow directly from Outlook. Extracts email context (subject, body, sender) and pre-fills a task creation form — with mock mode and real API support.

## Project Structure

```
proWorkFlow/
├── .eslintrc.json                       # ESLint config (Office Addins + React)
├── .hintrc                              # Webhint config
├── babel.config.json                    # Babel presets (env + TypeScript)
├── manifest.xml                         # Outlook add-in registration
├── webpack.config.js                    # Build config (multi-entry, HMR, HTTPS)
├── tsconfig.json                        # TypeScript config (ES5 target, strict)
├── package.json                         # Dependencies & scripts
│
├── .vscode/
│   ├── extensions.json                  # Recommended extensions
│   ├── launch.json                      # Debug profiles (Edge Chromium + Legacy)
│   ├── settings.json                    # Workspace settings
│   └── tasks.json                       # Build, debug, lint tasks
│
├── assets/                              # Icons & images
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   ├── icon-128.png
│   ├── logo-filled.png
│   └── mainLogo.png
│
├── src/
│   ├── commands/                        # Ribbon command functions (hidden page)
│   │   ├── commands.html
│   │   └── commands.ts
│   └── taskpane/                        # Main task pane UI (React SPA)
│       ├── index.tsx                    # React entry point (createRoot + FluentProvider)
│       ├── taskpane.html                # HTML shell with Office.js CDN
│       ├── taskpane.ts                  # Office.js utility functions
│       ├── hooks/
│       │   └── useEmailContext.ts       # Extract email data from Outlook item
│       ├── services/
│       │   ├── powerflowApi.ts          # Axios-based ProWorkflow API client (mock + real)
│       │   └── mockflow.ts             # Mock data for API responses (stub)
│       └── components/
│           ├── App.tsx                  # Root component (assembles Header + CreateTaskForm)
│           ├── Header.tsx               # Logo & title
│           ├── CreateTask/
│           │   └── CreateTaskForm.tsx   # Full task creation form
│           └── Layout/
│               └── AppLayout.tsx        # MUI drawer layout (for future routing)
│
└── docs/                                # Documentation (see below)
    ├── ARCHITECTURE.md
    ├── CODE_FLOW.md
    ├── DEBUGGING.md
    └── FILE_MAP.md
```

## Documentation

Detailed documentation for developers and LLMs:

| Document | Audience | Contents |
|---|---|---|
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | All | System architecture, component tree, tech stack, boot sequence, manifest structure |
| [docs/CODE_FLOW.md](docs/CODE_FLOW.md) | All | Complete execution paths, data flow diagrams, import/export graph, error handling |
| [docs/DEBUGGING.md](docs/DEBUGGING.md) | Developers | Debug setup, VS Code configs, troubleshooting, available scripts & tasks |
| [docs/FILE_MAP.md](docs/FILE_MAP.md) | LLMs | File-by-file reference with exports, imports, key lines, and dependency graph |

## Quick Start

```powershell
cd proWorkFlow
npm install
npm start            # Build + sideload + launch Outlook
```

Press **F5** in VS Code (with "Outlook Desktop (Edge Chromium)" launch config) to start debugging.

## Tech Stack

- **React 18.2** + **TypeScript 5.4** (ES5 target, strict mode)
- **Fluent UI React v9** (theming, components, CSS-in-JS)
- **MUI Material v9** (layout components — AppLayout drawer)
- **MUI Icons v9** (icon set)
- **Axios** (ProWorkflow API client, mock mode by default)
- **React Router DOM v7** (routing for future multi-view)
- **Emotion** (styling engine for MUI)
- **Office.js** (mailbox integration via CDN)
- **Webpack 5** (multi-entry bundling, HMR, HTTPS dev server)
- **core-js** + **regenerator-runtime** + **es6-promise** (IE11 polyfills)

## Development Scripts

| Script | Purpose |
|---|---|
| `npm start` | Build + sideload + launch Outlook desktop |
| `npm stop` | Stop debugging + remove sideload |
| `npm run dev-server` | Start dev server only (localhost:3000) |
| `npm run build` | Production build (webpack --mode production) |
| `npm run build:dev` | Development build with source maps |
| `npm run watch` | Rebuild on file changes |
| `npm run lint` | Run ESLint (office-addin-lint check) |
| `npm run lint:fix` | Auto-fix all fixable lint problems |
| `npm run prettier` | Format code with Prettier |
| `npm run validate` | Validate manifest.xml |
| `npm run signin` | M365 account login |
| `npm run signout` | M365 account logout |
