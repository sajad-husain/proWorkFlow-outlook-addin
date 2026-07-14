# proWorkFlow Outlook Add-in

Microsoft Office add-in for Outlook that creates and edits tasks in [ProWorkflow](https://proworkflow.com) directly from your inbox. Extracts email context (subject, body, sender, attachments) and pre-fills a task creation form via the ProWorkflow REST API v4.

---

## Features

- **Create Tasks from Emails** — Auto-fills task name from subject, builds description from body with sender/recipient metadata, attaches email attachments
- **Edit & Manage Tasks** — Browse, search, update status, reassign, and delete tasks per project
- **API Key Authentication** — First-run setup dialog that verifies and persists your ProWorkflow API key in localStorage
- **Draft Auto-Save** — Create form state persists to localStorage with debounce so you never lose work
- **Email Preview** — Collapsible email body preview with attachment listing inside the task pane
- **Keyboard Shortcuts** — Ctrl+Enter to submit, Escape to reset
- **Toast Notifications** — Reusable snackbar component for success/error feedback

---

## Project Architecture

```
proWorkFlow-outlook-addin/
│
├── .gitignore
├── README.md
│
└── proWorkFlow/                                  # Add-in root
    ├── .eslintrc.json                            # ESLint (office-addins/react plugin)
    ├── .hintrc                                   # Webhint config
    ├── apitest.html                              # Standalone browser-based API test page
    ├── babel.config.json                         # Babel presets (env + TypeScript)
    ├── manifest.xml                              # Outlook add-in manifest (MailApp, Mailbox 1.1+)
    ├── package.json                              # Scripts & dependencies
    ├── tsconfig.json                             # TypeScript (ES5 target, strict, JSX react)
    ├── webpack.config.js                         # Multi-entry webpack: taskpane + commands
    │
    ├── .vscode/
    │   ├── extensions.json                       # Recommended extensions
    │   ├── launch.json                           # Edge Chromium + Edge Legacy debug profiles
    │   ├── tasks.json                            # Build, dev-server, lint, start/stop tasks
    │   └── settings.json                         # ESLint validation settings
    │
    ├── assets/                                   # Add-in icons (16–128px) + logo
    │
    └── src/
        ├── commands/                             # Ribbon button command handler (hidden page)
        │   ├── commands.html                     # HTML shell loading office.js
        │   └── commands.ts                       # Registers `action` with Office.actions
        │
        └── taskpane/                             # Main task pane — React SPA
            ├── index.tsx                          # Entry: Office.onReady() → createRoot → <App/>
            ├── taskpane.html                      # HTML shell with Office.js CDN + IE11 fallback
            ├── taskpane.ts                        # Legacy entry (ReactDOM.render, unused)
            │
            ├── hooks/
            │   └── useLocalStorage.ts            # Custom hook for localStorage persistence
            │
            ├── services/
            │   ├── proworkflow.ts                 # ProWorkflow REST API client (CRUD)
            │   ├── outlook.ts                     # Outlook data extraction (subject, body, etc.)
            │   └── editTask.ts                    # Thin wrapper around proworkflow.ts
            │
            ├── theme/
            │   └── theme.ts                       # MUI theme (ProWorkflow blue palette)
            │
            └── components/
                ├── App.tsx                        # Root: API key gate + ThemeProvider
                ├── Header.tsx                     # AppBar with "New Task" / "Edit Task" tabs
                ├── Layout.tsx                     # Duplicate MUI theme definition
                ├── Toast.tsx                      # Reusable Snackbar + Alert notification
                │
                ├── ApiKeySetup/
                │   └── ApiKeySetup.tsx            # Dialog for entering/verifying API key
                │
                ├── CreateTask/
                │   └── CreateTask.tsx             # Full task creation form (~840 lines)
                │
                ├── EditTask/
                │   └── EditTask.tsx               # Task listing, editing, deletion (~840 lines)
                │
                ├── ErrorBoundary/
                │   └── ErrorBoundary.tsx          # React error boundary with reload fallback
                │
                └── Router/
                    └── AppRouter.tsx              # State-based router (create/edit views)
```

### Tech Stack

| Layer | Technology |
|---|---|
| Framework | React 18.2 + TypeScript 5.4 (ES5 target, strict mode) |
| UI Library | Material UI (MUI) v9 — theming, components, CSS-in-JS |
| Office Integration | Office.js (mailbox, item body, notifications via CDN) |
| API Backend | ProWorkflow REST API v4 (`https://api.proworkflow.com/api/v4`) |
| Bundler | Webpack 5 (multi-entry, HMR, HTTPS dev server on port 3000) |
| Transpilation | Babel (preset-env + TypeScript) + ts-loader |
| Polyfills | core-js, regenerator-runtime, es6-promise (IE11 compat) |
| Linting | ESLint (office-addins/react plugin) |
| Dev Tooling | office-addin-debugging, office-addin-dev-certs |

### Boot Sequence

1. **Office host** (Outlook) reads `manifest.xml` and loads `taskpane.html` into a webview.
2. **Office.js** loads from CDN (`appsforoffice.microsoft.com`) and initializes the host bridge.
3. **taskpane.html** renders `<div id="container">` with an IE11 compatibility fallback.
4. **index.tsx** calls `Office.onReady()` then `createRoot().render(<App />)`.
5. **App.tsx** checks for a stored API key in localStorage — if missing, shows `ApiKeySetup` dialog; otherwise renders the main `AppRouter`.
6. **AppRouter** renders either `CreateTask` or `EditTask` based on the selected header tab.

---

## Code Flow

### Task Pane (Main UI)

```
Outlook Ribbon → "Show Task Pane"
    │
    ▼
manifest.xml → <Action xsi:type="ShowTaskpane">
    │
    ▼
taskpane.html (https://localhost:3000/taskpane.html)
    │
    ▼
Office.js loads → Office.onReady()
    │
    ▼
index.tsx → createRoot → <App />
    │
    ▼
App.tsx
    ├── No API key? → <ApiKeySetup /> (verify + store key)
    └── API key exists? → <AppRouter />
         ├── "New Task" tab → <CreateTask />
         │    ├── Extract email context (subject, body, sender, attachments)
         │    ├── Fetch projects + contacts from ProWorkflow API
         │    ├── Pre-fill form from email data
         │    └── POST /tasks on submit
         │
         └── "Edit Task" tab → <EditTask />
              ├── Fetch projects → fetch tasks per project
              ├── Search/filter task list
              ├── Edit form → PUT /tasks/:id
              └── Delete → DELETE /tasks/:id
```

### Commands (Ribbon Action)

```
Outlook Ribbon → "Perform an action"
    │
    ▼
manifest.xml → <Action xsi:type="ExecuteFunction">
    │                    <FunctionName>action</FunctionName>
    ▼
commands.html → commands.ts → Office.actions.associate("action", ...)
    │
    ▼
Shows notification: "Performed action." → event.completed()
```

### Services Layer

```
proworkflow.ts          outlook.ts              editTask.ts
┌────────────────┐     ┌─────────────────────┐  ┌──────────────┐
│ testApiKey()   │     │ getOutlookItemData() │  │ Thin wrapper │
│ getProjects()  │     │ getOutlookItemDataAsync()│ │ around       │
│ getUsers()     │     │ getAttachmentContent()│  │ proworkflow  │
│ createTask()   │     │ cleanEmailBody()     │  │ .ts methods  │
│ getTasks()     │     │ generateTaskDesc()   │  └──────────────┘
│ getTask()      │     └─────────────────────┘
│ updateTask()   │
│ deleteTask()   │
└────────────────┘
```

---

## Getting Started

### Prerequisites

- **Node.js 18+**
- **Outlook desktop client** (Windows/Mac) with M365, or Outlook on the web
- A **Microsoft 365 account** for sideloading — free dev sandbox available at [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- A **ProWorkflow account** with an API key

### Install & Run

```powershell
cd proWorkFlow
npm install
npm start
```

This builds the project, generates dev certificates, sideloads the add-in into Outlook, and launches Outlook desktop.

### First-Time Setup

1. After sideloading, the task pane opens in Outlook.
2. Enter your **ProWorkflow API key** in the setup dialog.
3. The key is verified and stored in localStorage.
4. Start creating tasks from your emails.

---

## Development

### VS Code Debugging (Recommended)

1. Open the `proWorkFlow` folder in VS Code (not the repo root).
2. Press **F5** → select **"Outlook Desktop (Edge Chromium)"**.
3. VS Code builds, starts dev server, sideloads the add-in, and attaches the debugger on port **9229**.

### Launch Configurations

| Profile | WebView | Port | Notes |
|---|---|---|---|
| Outlook Desktop (Edge Chromium) | Edge Chromium | 9229 | Default; works with latest Office |
| Outlook Desktop (Edge Legacy) | Edge Legacy | 9222 | For older Office versions |

### Manual Dev Server

```powershell
npm run dev-server        # Starts webpack-dev-server on https://localhost:3000
```

Then sideload manually:
```powershell
npx office-addin-debugging start manifest.xml --app outlook
```

### Debugging Tips

- **Source maps** enabled in dev mode — set breakpoints directly in `.ts`/`.tsx` files.
- **Console logging** — use `console.log()` and view in the debugger console.
- **HMR** — React component changes reflect without full page reload.
- **Attach to running instance** — if Outlook is already open, use the "attach" config on port 9229.

---

## Scripts Reference

| Script | Description |
|---|---|
| `npm start` | Build + sideload + launch Outlook desktop |
| `npm stop` | Stop debugging + remove sideload |
| `npm run dev-server` | Start dev server only (localhost:3000) |
| `npm run build` | Production build (webpack --mode production) |
| `npm run build:dev` | Development build with source maps |
| `npm run watch` | Rebuild on file changes |
| `npm run lint` | Run ESLint |
| `npm run lint:fix` | Auto-fix lint issues |
| `npm run prettier` | Format code with Prettier |
| `npm run validate` | Validate manifest.xml structure |
| `npm run signin` | M365 account login for sideloading |
| `npm run signout` | M365 account logout |

### VS Code Tasks

| Task | Purpose |
|---|---|
| Build (Development) | `npm run build:dev` |
| Build (Production) | `npm run build` |
| Dev Server | `npm run dev-server` |
| Debug: Outlook Desktop | `npm start` |
| Stop Debug | `npm stop` |
| Watch | `npm run watch` |
| Lint: Check | `npm run lint` |
| Lint: Fix | `npm run lint:fix` |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `office.js` not loaded | Check network — office.js loads from CDN, requires internet |
| Add-in not appearing | Run `npm start` to sideload; validate manifest with `npm run validate` |
| HTTPS certificate warning | Accept the self-signed dev cert in your browser |
| "This add-in will not run" | IE/Edge Legacy webview detected — use Outlook with Edge Chromium (Office 2021+ or M365) |
| Debugger not attaching | Close all Edge/Chrome instances; ensure port 9229 is free |
| Changes not reflecting | Try `npm run build:dev` then refresh the task pane |
| API key rejected | Verify your ProWorkflow API key is valid at [proworkflow.com](https://proworkflow.com) |
| Tasks not loading | Check project selection — tasks are fetched per-project from the API |

---

## License

MIT
