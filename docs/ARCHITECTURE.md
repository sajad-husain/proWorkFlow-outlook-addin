# Architecture

## Overview

proWorkFlow is an Outlook add-in (MailApp) built with React 18, TypeScript, Fluent UI v9, and the Office.js SDK. It executes inside an Outlook task pane, runs in a browser-based webview (Edge Chromium), and communicates with the Outlook mailbox via the Office.js API bridge.

## High-Level Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    Outlook Desktop                        │
│  ┌──────────────────────────────────────────────────┐   │
│  │              WebView (Edge Chromium)              │   │
│  │  ┌──────────────────────────────────────┐       │   │
│  │  │   taskpane.html (React SPA)          │       │   │
│  │  │   ┌────────────────────────────┐     │       │   │
│  │  │   │  FluentProvider (theme)    │     │       │   │
│  │  │   │  ┌────────────────────┐   │     │       │   │
│  │  │   │  │  App               │   │     │       │   │
│  │  │   │  │  ├─ Header         │   │     │       │   │
│  │  │   │  │  └─ CreateTaskForm │   │     │       │   │
│  │  │   │  └────────────────────┘   │     │       │   │
│  │  │   │  ┌────────────────────┐   │     │       │   │
│  │  │   │  │  useEmailContext   │   │     │       │   │
│  │  │   │  │  (hook)            │   │     │       │   │
│  │  │   │  └────────────────────┘   │     │       │   │
│  │  │   └────────────────────────────┘     │       │   │
│  │  │                                              │   │
│  │  │  Services:                                   │   │
│  │  │  ┌──────────────────────┐                    │   │
│  │  │  │  powerflowApi.ts     │── Axios ──▶ PWF    │   │
│  │  │  │  (mock / real)       │     API            │   │
│  │  │  └──────────────────────┘                    │   │
│  │  └──────────────────────────────────────────────┘   │
│  │                                                   │   │
│  │  ┌──────────────────────────────────────┐       │   │
│  │  │   commands.html (hidden page)        │       │   │
│  │  │   └─ action() function              │       │   │
│  │  └──────────────────────────────────────┘       │   │
│  └──────────────────────────────────────────────────┘   │
│                                                         │
│  Office.js API Bridge                                   │
│  ┌──────────────────────────────────────────────────┐   │
│  │  Office.context.mailbox.item.*                    │   │
│  │  ├─ item.subject (sync)                          │   │
│  │  ├─ item.from.emailAddress (sync)                │   │
│  │  ├─ item.body.getAsync() (async)                 │   │
│  │  ├─ item.attachments (sync)                      │   │
│  │  └─ item.itemId (sync)                           │   │
│  └──────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────┘
```

## Project Structure

```
proWorkFlow/
├── .eslintrc.json                   # ESLint config (Office Addins + React)
├── .hintrc                          # Webhint config
├── babel.config.json                # Babel presets (env + TypeScript)
├── manifest.xml                     # Add-in manifest (Outlook registration)
├── webpack.config.js                # Build configuration (multi-entry)
├── tsconfig.json                    # TypeScript configuration
├── package.json                     # Dependencies and scripts
│
├── .vscode/
│   ├── extensions.json              # Recommended extensions
│   ├── launch.json                  # Debug launch configurations
│   ├── settings.json                # Workspace settings
│   └── tasks.json                   # Build & debug tasks
│
├── assets/                          # Icons and images
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   ├── icon-128.png
│   ├── logo-filled.png
│   └── mainLogo.png
│
├── src/
│   ├── commands/
│   │   ├── commands.html           # Hidden HTML shell for ribbon command functions
│   │   └── commands.ts             # Ribbon action handler (ExecuteFunction)
│   └── taskpane/
│       ├── index.tsx               # React entry point (bootstraps the app)
│       ├── taskpane.html           # HTML shell loaded in the task pane
│       ├── taskpane.ts             # Office.js utility functions
│       ├── hooks/
│       │   └── useEmailContext.ts  # Hook to extract email data from Outlook item
│       ├── services/
│       │   ├── powerflowApi.ts     # Axios-based ProWorkflow API client (mock + real)
│       │   └── mockflow.ts        # Mock API response data (stub)
│       └── components/
│           ├── App.tsx             # Root React component
│           ├── Header.tsx          # Logo + title header
│           ├── CreateTask/
│           │   └── CreateTaskForm.tsx  # Full task creation form
│           └── Layout/
│               └── AppLayout.tsx   # MUI drawer layout (for future routing)
│
└── docs/                           # Documentation
    ├── ARCHITECTURE.md
    ├── CODE_FLOW.md
    ├── DEBUGGING.md
    └── FILE_MAP.md
```

## Entry Points

The webpack build produces two separate HTML entry points:

| Entry | Chunks | Purpose |
|---|---|---|
| `taskpane.html` | `polyfill` + `react` + `taskpane` | Main UI loaded in the task pane |
| `commands.html` | `polyfill` + `commands` | Hidden page for ribbon command execution |

### Build Splitting (webpack.config.js:20-28)

```javascript
entry: {
  polyfill: ["core-js/stable", "regenerator-runtime/runtime"],  // IE11 polyfills
  react: ["react", "react-dom"],                                  // Vendor chunk
  taskpane: { import: ["./src/taskpane/index.tsx", ...], dependOn: "react" },
  commands: "./src/commands/commands.ts"
}
```

The `react` chunk is separated to enable Hot Module Replacement in the task pane. The `polyfill` chunk provides ES6+ compatibility for legacy Office webviews (IE/Edge Legacy).

## Boot Sequence

```
Outlook loads add-in (manifest.xml) — opens on message read
         │
         ▼
Ribbon appears in read message view
         │
         ├── [User clicks "Create Task" ribbon button]
         │         │
         │         ▼
         │   taskpane.html loads in task pane
         │         │
         │         ▼
         │   office.js loaded from CDN
         │         │
         │         ▼
         │   index.tsx calls Office.onReady()
         │         │
         │         ▼
         │   React root created (createRoot)
         │         │
         │         ▼
          │   <FluentProvider theme={webLightTheme}>
          │     <App />
          │   </FluentProvider>
          │         │
          │         ▼
          │   App renders:
          │   ├─ <Header logo="assets/logo-filled.png" />
          │   └─ <CreateTaskForm emailData emailData loading loading />
         │         │
         │         ▼
         │   useEmailContext hook fires:
         │   ├─ Reads item.subject (sync)
         │   ├─ Reads item.from (sync)
         │   ├─ Calls item.body.getAsync() (async)
         │   ├─ Reads item.attachments (sync)
         │   └─ Returns emailData to CreateTaskForm
         │         │
         │         ▼
         │   Form pre-fills title & description
         │   from email subject & body
         │
         └── [User clicks "Perform an action"]
                   │
                   ▼
             commands.html loads (hidden)
                   │
                   ▼
             commands.ts: action() executes
                   │
                   ▼
             Notification shown in email
                   │
                   ▼
             event.completed() called
```

## Component Tree

```
<FluentProvider theme={webLightTheme}>              [index.tsx:19]
  └── <App>                                         [App.tsx:16]
        ├── <Header                                 [Header.tsx]
        │     logo="assets/logo-filled.png" />
        └── <CreateTaskForm                         [CreateTaskForm.tsx]
              emailData={emailData}
              loading={loading} />

  Hook: useEmailContext()                            [hooks/useEmailContext.ts]
    ├── Office.context.mailbox.item.subject (sync)
    ├── Office.context.mailbox.item.from (sync)
    ├── Office.context.mailbox.item.body.getAsync() (async)
    ├── Office.context.mailbox.item.attachments (sync)
    └── Returns { emailData, loading }
```

## Technology Stack

| Layer | Technology | Purpose |
|---|---|---|
| UI Framework | React 18.2 | Component-based UI |
| Design System | Fluent UI React v9 | Theming, components, styling |
| Layout Components | MUI Material v9 | Drawer navigation, AppBar |
| Icons | Fluent UI React Icons v2 + MUI Icons | Icon set |
| Language | TypeScript 5.4 (target ES5) | Type safety, IE11 compatibility |
| Build | Webpack 5 | Module bundling, HMR |
| Transpilation | Babel + ts-loader | TypeScript → ES5 |
| Office API | Office.js | Outlook mailbox integration |
| API Client | Axios 1.18 | ProWorkflow REST API (mock mode by default) |
| Routing | React Router DOM v7 | Multi-view navigation (future) |
| Styling | `makeStyles` (Fluent UI CSS-in-JS) + MUI `sx` | Component styling |
| Polyfills | core-js, regenerator-runtime, es6-promise | Legacy browser support |
| Dev Server | webpack-dev-server 5.1 | HTTPS local dev with HMR |
| Dev Certs | office-addin-dev-certs | Auto-generated HTTPS certificates |
| Linting | ESLint (office-addins plugin) | Code quality |
| Debugging | office-addin-debugging | Sideload + launch for testing |

## State Management

Currently uses **local React state** via `useState` + a custom `useEmailContext` hook:

| State | Location | Type | Default |
|---|---|---|---|
| `emailData` | hooks/useEmailContext.ts:13 | `useState<EmailData \| null>` | `null` |
| `loading` | hooks/useEmailContext.ts:14 | `useState<boolean>` | `true` |
| `formData` (workspace, taskList, title, etc.) | CreateTaskForm.tsx:148 | `useState<object>` | empty fields |
| `isSubmitting` | CreateTaskForm.tsx:160 | `useState<boolean>` | `false` |

No global state management library is used. The app is a single-view SPA (future routing with React Router planned).

## Manifest Configuration

The add-in is registered with Outlook via `manifest.xml`:

- **Type:** `MailApp` (Outlook add-in)
- **Host:** `Mailbox`
- **Permissions:** `ReadWriteItem`
- **Activation:** Message Compose form (ribbon buttons appear when composing an email)
- **Minimum Mailbox API:** v1.1 (v1.3 for version overrides)

### Ribbon Configuration

```
Message Compose Command Surface
  └── Group: "Contoso Add-in"
        ├── Button "Show Task Pane"  →  ShowTaskpane  →  taskpane.html
        └── Button "Perform an action"  →  ExecuteFunction  →  action()
```

## Key Design Decisions

1. **ES5 target + polyfills** for compatibility with older Office webviews (IE/Edge Legacy used in older Outlook desktop versions).

2. **Chunk splitting** (`polyfill`, `react`, `taskpane`, `commands`) optimizes loading and enables HMR for the React app without affecting command functions.

3. **CSS-in-JS via Fluent UI `makeStyles`** rather than separate CSS files. Styles use Fluent UI design tokens for consistent theming.

4. **Office.js loaded from CDN** (not bundled) since it is provided by the Office host at runtime.

5. **Separate entry points** for taskpane and commands because they have different lifecycle requirements: the taskpane is a persistent UI, while commands are transient function executions.

6. **Mock-first API layer** — `powerflowApi.ts` defaults to mock mode so the UI is fully functional without real API credentials. Flip `USE_MOCK = false` when ready for production.

7. **Email context extraction via hook** — `useEmailContext` abstracts Office.js mailbox API calls into a reusable hook with a test fallback for development outside Outlook.
