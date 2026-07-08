# proWorkFlow Outlook Add-in

Outlook add-in that lets you create tasks in ProWorkflow directly from Outlook. Extracts email context (subject, body, sender) and pre-fills a task creation form — with mock mode and real API support.

---

## Project Architecture

```
proWorkFlow-outlook-addin/
│
├── .gitignore
├── README.md
│
└── proWorkFlow/                              # Add-in root
    ├── .eslintrc.json                        # ESLint config (Office Addins + React)
    ├── .hintrc                               # Webhint config
    ├── babel.config.json                     # Babel presets (env + TypeScript)
    ├── manifest.xml                          # Outlook add-in manifest (MailApp, Mailbox 1.1+)
    ├── webpack.config.js                     # Multi-entry webpack: taskpane + commands
    ├── tsconfig.json                         # TypeScript (ES5 target, strict, JSX react)
    ├── package.json                          # Scripts & dependencies
    │
    ├── .vscode/
    │   ├── extensions.json                   # Recommended extensions
    │   ├── launch.json                       # Edge Chromium + Edge Legacy debug profiles
    │   ├── tasks.json                        # Build, dev-server, lint, start/stop debug tasks
    │   └── settings.json                     # Workspace settings
    │
    ├── assets/                               # Icons (16–128px) + logo
    │
    └── src/
        ├── commands/                         # Ribbon button command handler (hidden page)
        │   ├── commands.html                 # HTML shell loading office.js
        │   └── commands.ts                   # Register `action` function with Office.actions
        │
        └── taskpane/                         # Main task pane — React SPA
            ├── taskpane.html                 # HTML shell with Office.js CDN + IE11 fallback
            ├── taskpane.ts                   # Utility: insertText() into email body
            ├── index.tsx                     # React entry: createRoot + FluentProvider
            └── components/
                ├── App.tsx                   # Root component
                ├── Header.tsx                # Header placeholder
                └── Router/
                    └── AppRouter.tsx         # Router placeholder (empty)
```

### Tech Stack

| Layer | Technology |
|---|---|
| Framework | React 18.2 + TypeScript 5.4 (ES5 target, strict mode) |
| UI Library | Fluent UI React v9 (theming, components, CSS-in-JS) |
| Office Integration | Office.js (mailbox, item body, notifications via CDN) |
| Bundler | Webpack 5 (multi-entry, HMR, HTTPS dev server) |
| Transpilation | Babel (preset-env + TypeScript) + ts-loader |
| Polyfills | core-js, regenerator-runtime, es6-promise (IE11) |
| Linting | ESLint (office-addins/react plugin) |
| Dev Tooling | office-addin-debugging, office-addin-dev-certs |

### Boot Sequence

1. **Office host** (Outlook) reads `manifest.xml` and loads `taskpane.html` or `commands.html` into a webview.
2. **Office.js** loads from CDN (`appsforoffice.microsoft.com`) and initializes the host bridge.
3. **taskpane.html** renders `<div id="container">` and shows an IE11 compatibility fallback if needed.
4. **index.tsx** calls `Office.onReady()` → then `createRoot().render(<App />)`.
5. **App.tsx** (currently a placeholder) renders `<h2>Apps</h2>` into the FluentUI theme context.

---

## Code Flow

### Task Pane Path (UI Entry Point)

```
Outlook Ribbon Button Click
    │
    ▼
manifest.xml → <Action xsi:type="ShowTaskpane">
    │                    resid="Taskpane.Url"
    ▼
taskpane.html (https://localhost:3000/taskpane.html)
    │
    ▼
office.js loads & initializes → Office.onReady()
    │
    ▼
index.tsx: createRoot → <App />
    │
    ▼
App.tsx:  <h2>Apps</h2>     // Placeholder for future form
```

### Commands Path (Ribbon Button Action)

```
Outlook Ribbon "Perform an action" Click
    │
    ▼
manifest.xml → <Action xsi:type="ExecuteFunction">
    │                    <FunctionName>action</FunctionName>
    ▼
commands.html (https://localhost:3000/commands.html)
    │
    ▼
commands.ts: Office.onReady() → action(event)
    │
    ▼
Shows notification: "Performed action."
    │
    ▼
event.completed()   // Required to signal async completion
```

### Data Flow (Future / Planned)

```
Outlook Mail Item
    │  (Office.context.mailbox.item)
    ▼
useEmailContext hook  →  extracts subject, body, sender
    │
    ▼
CreateTaskForm.tsx    →  pre-fills task fields
    │
    ▼
powerflowApi.ts      →  POST /tasks (Axios)
    │                    mock mode: returns mockflow.ts data
    ▼
ProWorkflow API (or mock response)
```

### Import Dependency Graph (Current)

```
index.tsx
  ├── Office.onReady() (global)
  └── App.tsx
       └── (React, FluentProvider - via index.tsx)

commands.ts
  └── Office.actions, Office.context.mailbox.item

taskpane.ts
  └── Office.context.mailbox.item.body.setSelectedDataAsync
```

---

## Debugging

### Prerequisites

- Node.js 18+
- Outlook desktop client (or Outlook on the web)
- A Microsoft 365 account (for sideloading) — or use the free dev sandbox from [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)

### Quick Start Debug

```powershell
cd proWorkFlow
npm install
npm start
```

This builds the project, generates dev certificates, sideloads the add-in into Outlook, and launches Outlook desktop.

### VS Code Debugging (Recommended)

1. Open the `proWorkFlow` folder (not the repo root) in VS Code — this ensures `.vscode/launch.json` is active.
2. Press **F5** (or Run → Start Debugging).
3. Select **"Outlook Desktop (Edge Chromium)"** from the dropdown.
4. VS Code will:
   - Run the **"Debug: Outlook Desktop"** pre-launch task (builds + starts dev server + sideloads)
   - Attach the debugger to the Edge Chromium webview on port **9229**
   - Open Outlook with the add-in loaded

### Available Launch Configurations

| Profile | WebView | Port | Notes |
|---|---|---|---|
| Outlook Desktop (Edge Chromium) | Edge Chromium | 9229 | Default; works with latest Office |
| Outlook Desktop (Edge Legacy) | Edge Legacy | 9222 | For older Office versions |

Both profiles run the `Debug: Outlook Desktop` pre-launch task and `Stop Debug` post-debug task.

### Manual Dev Server

```powershell
npm run dev-server     # Starts webpack-dev-server on https://localhost:3000
```

Then sideload manually via:
```powershell
npx office-addin-debugging start manifest.xml --app outlook
```

### Debugging Techniques

- **Source maps** are enabled in dev mode (`devtool: "source-map"` in webpack) — set breakpoints directly in `.ts`/`.tsx` files.
- **Console logging** — Office.js add-ins run in a webview; use `console.log()` and view output in the debugger console.
- **Attach to running instance** — If Outlook is already open with the add-in loaded, use the "attach" request in `launch.json` (port 9229).
- **HMR (Hot Module Replacement)** — The webpack dev server supports HMR; changes to React components reflect without full reload.

### VS Code Tasks

| Task | Command | Purpose |
|---|---|---|
| Build (Development) | `npm run build:dev` | Production build with source maps |
| Build (Production) | `npm run build` | Minified production build |
| Dev Server | `npm run dev-server` | Start webpack-dev-server only |
| Debug: Outlook Desktop | `npm start -- desktop --app outlook` | Build + sideload + launch Outlook |
| Stop Debug | `npm stop` | Remove sideload + clean up |
| Watch | `npm run watch` | Rebuild on file changes |
| Lint: Check | `npm run lint` | Run ESLint |
| Lint: Fix | `npm run lint:fix` | Auto-fix lint issues |

### Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| `office.js` not loaded | Check network tab in devtools; office.js loads from CDN — requires internet |
| Add-in not appearing in Outlook | Run `npm start` to sideload; check `manifest.xml` is valid with `npm run validate` |
| HTTPS certificate warning | Dev certs auto-generated; accept the self-signed cert in browser |
| "This add-in will not run" message | IE/Edge Legacy webview detected; use Outlook with Edge Chromium (Office 2021+ or M365) |
| Debugger not attaching | Close all Edge/Chrome instances; ensure port 9229 is free |
| Changes not reflecting | HMR may not pick up all changes — try `npm run build:dev` then refresh the task pane |

### Scripts Reference

| Script | Description |
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
| `npm run validate` | Validate manifest.xml structure |
| `npm run signin` | M365 account login for sideloading |
| `npm run signout` | M365 account logout |
