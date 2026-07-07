# proWorkFlow Outlook Add-in

Outlook Classic add-in that lets you create tasks in ProWorkflow directly from Outlook.

> **Status:** Template scaffold. Currently contains the default Office Add-in Task Pane React template code. ProWorkflow API integration is not yet implemented.

## Project Structure

The codebase lives in `proWorkFlow/` and follows the standard Office Add-in Task Pane + React template structure:

```
proWorkFlow/
├── src/
│   ├── commands/                    # Ribbon command functions (hidden page)
│   │   ├── commands.html
│   │   └── commands.ts
│   └── taskpane/                    # Main task pane UI (React SPA)
│       ├── index.tsx                # React entry point
│       ├── taskpane.html            # HTML shell
│       ├── taskpane.ts              # Office.js integration
│       └── components/
│           ├── App.tsx              # Root component
│           ├── Header.tsx           # Logo & welcome
│           ├── HeroList.tsx         # Feature list
│           └── TextInsertion.tsx    # Text input & insert
├── manifest.xml                     # Outlook add-in registration
├── webpack.config.js                # Build config (multi-entry)
├── tsconfig.json                    # TypeScript config
├── package.json                     # Dependencies & scripts
├── assets/                          # Icons & images
└── docs/                            # Documentation (see below)
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

## Branching / Workflow

This is a template project. All source code is in the `proWorkFlow/` subdirectory. The root `README.md` and `docs/` directory are at the repository root level.
