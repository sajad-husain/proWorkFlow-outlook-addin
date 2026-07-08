# File Map — LLM Reference

This document provides a complete, structured map of every source file in the project. It is designed for quick lookup by LLMs and developers, showing file relationships, exports, imports, and key code locations.

---

## Source Files

### `src/taskpane/index.tsx`
**Purpose:** React application entry point. Boots the task pane UI.

| Detail | Value |
|---|---|
| **Imports** | `react`, `react-dom/client` (`createRoot`), `./components/App`, `@fluentui/react-components` (`FluentProvider`, `webLightTheme`) |
| **Exports** | None (side-effect module) |
| **Entry** | `Office.onReady(() => root.render(<FluentProvider><App/></FluentProvider>))` |
| **Lines** | 36 |
| **Key line** | `19: Office.onReady(...)` — app boot after Office.js init |
| **HMR** | Lines 21-30: `module.hot.accept` |

---

### `src/taskpane/taskpane.ts`
**Purpose:** Office.js utility functions.

| Detail | Value |
|---|---|
| **Imports** | None (uses global `Office` and `console`) |
| **Exports** | `insertText(text: string): Promise<void>` |
| **Lines** | 18 |
| **Key line** | `6: Office.context.mailbox.item?.body.setSelectedDataAsync(...)` |
| **Error handling** | Lines 10-12: throws on `AsyncResultStatus.Failed` |
| **Status** | Utility file — `insertText` is not currently imported by any component |

---

### `src/taskpane/components/App.tsx`
**Purpose:** Root React component. Assembles the UI from child components.

| Detail | Value |
|---|---|
| **Imports** | `Header`, `CreateTaskForm`, `useEmailContext`, `@fluentui/react-components` (`makeStyles`, `tokens`) |
| **Exports** | `App` |
| **Props** | None (stateless functional component) |
| **Lines** | 28 |
| **Key line** | `18: const { emailData, loading } = useEmailContext()` — extracts email data via hook |
| **Key line** | `22: <Header logo="assets/logo-filled.png" />` — renders header with logo |
| **Key line** | `23: <CreateTaskForm emailData={emailData} loading={loading} />` — passes data to form |

---

### `src/taskpane/components/Header.tsx`
**Purpose:** ⚠️ Currently contains a duplicate of the old App component (with `AppProps`, `title` prop, `mainLogo.png`). Needs cleanup to be a proper Header component.

| Detail | Value |
|---|---|
| **Imports** | `Header` (self-import), `CreateTaskForm`, `useEmailContext`, `@fluentui/react-components` (`makeStyles`, `tokens`) |
| **Exports** | `App` and `AppProps` (not Header) |
| **Props** | `{ title: string }` |
| **Lines** | 32 |
| **Status** | Orphan — exports App but should export a Header component; not independently used |

---

### `src/taskpane/components/CreateTask/CreateTaskForm.tsx`
**Purpose:** Full task creation form with workspace selection, task list, title, description, assignee, due date, urgency, and attachment options.

| Detail | Value |
|---|---|
| **Imports** | `react` (`useState`, `useEffect`), `@fluentui/react-components` (varied), `@fluentui/react-icons` |
| **Exports** | `CreateTaskForm(props: CreateTaskFormProps)` |
| **Props** | `{ emailData: EmailData \| null, loading: boolean }` |
| **States** | `formData` (workspace, taskList, title, etc.), `isSubmitting` |
| **Lines** | 418 |
| **Key line** | `163: useEffect` — pre-fills title/description from emailData |
| **Key line** | `191: handleSubmit` — simulates task creation (dispatches toast) |
| **Notifications** | Uses `useToastController` for success/error toasts |

---

### `src/taskpane/hooks/useEmailContext.ts`
**Purpose:** Custom hook that extracts email context from the current Outlook item.

| Detail | Value |
|---|---|
| **Imports** | `react` (`useState`, `useEffect`) |
| **Exports** | `useEmailContext()` (function), `EmailData` (interface) |
| **Returns** | `{ emailData: EmailData \| null, loading: boolean }` |
| **Lines** | 70 |
| **Key line** | `19: const item = Office.context.mailbox.item` — accesses Outlook item |
| **Key line** | `30: item.body?.getAsync('text', ...)` — async body extraction |
| **Fallback** | Lines 58-65: mock data for testing outside Outlook |

---

### `src/taskpane/services/powerflowApi.ts`
**Purpose:** ProWorkflow API client with mock mode and real Axios-based HTTP client.

| Detail | Value |
|---|---|
| **Imports** | `axios`, from `./mockData` (`mockProjects`, `mockAssignees`, `mockTaskCreationResponse`) — ⚠️ file is `mockflow.ts` |
| **Exports** | `proWorkflowApi` (object with `getProjects`, `getAssignees`, `createTask`), `setApiMode` |
| **Mock mode** | `USE_MOCK = true` by default |
| **Key line** | `28: getProjects` — fetches projects (mock or API) |
| **Key line** | `56: createTask` — creates a task (mock or API) |
| **Lines** | 75 |

---

### `src/taskpane/services/mockflow.ts`
**Purpose:** Mock data for API responses during development.

| Detail | Value |
|---|---|
| **Exports** | `mockProjects`, `mockAssignees`, `mockTaskCreationResponse` |
| **Used by** | `powerflowApi.ts` |
| **Status** | Empty stub — needs mock data arrays |
| **⚠️ Note** | `powerflowApi.ts` imports from `./mockData` but the file on disk is `mockflow.ts` — needs alignment |

---

### `src/taskpane/components/Layout/AppLayout.tsx`
**Purpose:** MUI-based drawer layout with navigation for future multi-view routing.

| Detail | Value |
|---|---|
| **Imports** | `react` (`useState`), `@mui/material` (`AppBar`, `Toolbar`, `Drawer`, etc.), `@mui/icons-material`, `react-router-dom` (`useNavigate`) |
| **Exports** | `AppLayout({ children, title })` |
| **Menu items** | "Create Task" → `/create-task`, "Edit Task" → `/edit-task` |
| **Lines** | 134 |
| **Key line** | `24: AppLayout` — responsive drawer with permanent/temporary variants |
| **Status** | Not yet wired into the app boot flow (standalone component) |

---

<!-- HeroList.tsx and TextInsertion.tsx were removed in customization (legacy template components) -->

### `src/commands/commands.ts`
**Purpose:** Handles ribbon command button actions (hidden page execution).

| Detail | Value |
|---|---|
| **Imports** | None (uses global `Office`) |
| **Exports** | None (side-effect module) |
| **Functions** | `action(event: Office.AddinCommands.Event)` |
| **Registration** | `Office.actions.associate("action", action)` at line 35 |
| **Lines** | 35 |
| **Key line** | `16: function action(event)` — entry point triggered by ribbon button |
| **Key line** | `25: notificationMessages.replaceAsync(...)` — shows notification |
| **Key line** | `31: event.completed()` — signals command completion |

---

### `src/taskpane/taskpane.html`
**Purpose:** HTML shell for the task pane.

| Detail | Value |
|---|---|
| **CDN loads** | `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` |
| **Mount point** | `<div id="container">` |
| **Fallback** | `<div id="tridentmessage">` for IE/Edge Legacy detection |
| **Lines** | 26 |

---

### `src/commands/commands.html`
**Purpose:** Hidden HTML page for ribbon command execution.

| Detail | Value |
|---|---|
| **CDN loads** | `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` |
| **Body** | Empty (no UI) |
| **Lines** | 15 |

---

## Configuration Files

### `manifest.xml`
**Purpose:** Registers the add-in with Outlook.

| Detail | Value |
|---|---|
| **Type** | `MailApp` |
| **Id** | `05c2e1c9-3e1d-406e-9a91-e9ac64854143` |
| **DisplayName** | `ProWorkFlow` |
| **Permissions** | `ReadWriteItem` |
| **Requirement** | Mailbox API v1.1 (v1.3 for version overrides) |
| **Activation** | Message Compose form |
| **Ribbon** | `Show Task Pane` → `ShowTaskpane` → taskpane.html, `Perform an action` → `ExecuteFunction` → `action()` |
| **AppDomains** | `https://www.contoso.com` |
| **Lines** | 121 |

---

### `webpack.config.js`
**Purpose:** Build configuration for both development and production.

| Detail | Value |
|---|---|
| **Dev URL** | `https://localhost:3000/` |
| **Prod URL** | `https://www.contoso.com/` (placeholder) |
| **Entries** | `polyfill`, `react`, `taskpane`, `commands` |
| **Plugins** | HtmlWebpackPlugin (×2), CopyWebpackPlugin, ProvidePlugin |
| **DevServer** | Port 3000, HTTPS, HMR, CORS enabled |
| **Lines** | 111 |

---

### `tsconfig.json`
**Purpose:** TypeScript compiler configuration.

| Detail | Value |
|---|---|
| **Target** | `es5` (IE11 compatibility) |
| **Module** | `ES2020` with `node` resolution |
| **JSX** | `react` |
| **Source maps** | `true` |
| **Strict mode** | Not enabled (only `noImplicitReturns`, `noUnusedParameters`) |
| **Lines** | 37 |

---

### `package.json`
**Purpose:** Project metadata, dependencies, and scripts.

| Detail | Value |
|---|---|
| **Name** | `office-addin-taskpane-react` |
| **app_to_debug** | `outlook` |
| **app_type_to_debug** | `desktop` |
| **dev_server_port** | `3000` |
| **Scripts** | 12 scripts (build, dev-server, start, stop, lint, etc.) |
| **Dependencies** | React 18, Fluent UI v9, MUI v9, Axios, React Router DOM v7, Emotion |
| **DevDependencies** | TypeScript 5.4, Webpack 5, Babel, ts-loader, office-addin-* tools |
| **Lines** | 82 |

---

### `.vscode/launch.json`
**Purpose:** VS Code debug launch configurations.

| Detail | Value |
|---|---|
| **Profile 1** | `Outlook Desktop (Edge Chromium)` — port 9229 |
| **Profile 2** | `Outlook Desktop (Edge Legacy)` — port 9222 |
| **PreLaunchTask** | `Debug: Outlook Desktop` |
| **PostDebugTask** | `Stop Debug` |
| **Lines** | 26 |

---

### `.vscode/tasks.json`
**Purpose:** VS Code build and debug tasks.

| Detail | Value |
|---|---|
| **Tasks** | Build (2), Debug (4), Dev Server, Install, Lint (2), Stop Debug, Watch |
| **Total tasks** | 11 |
| **Lines** | 160 |

---

### `.eslintrc.json`
**Purpose:** ESLint configuration.

| Detail | Value |
|---|---|
| **Plugin** | `office-addins` |
| **Extends** | `plugin:office-addins/react` |

---

## Assets

| File | Size (px) | Purpose |
|---|---|---|
| `assets/icon-16.png` | 16×16 | Ribbon icon (small) |
| `assets/icon-32.png` | 32×32 | Ribbon icon (medium) |
| `assets/icon-64.png` | 64×64 | Icon URL (manifest) |
| `assets/icon-80.png` | 80×80 | Notification icon |
| `assets/icon-128.png` | 128×128 | High-resolution icon URL |
| `assets/logo-filled.png` | — | Header logo in task pane |
| `assets/mainLogo.png` | — | Secondary logo |

---

## Dependency Graph (npm packages)

```
react 18.2 ─────────────────────────────────┐
react-dom 18.2 ─────────────────────────────┤
@fluentui/react-components 9.55.1 ──────────┤
@fluentui/react-icons 2.0.264 ──────────────┤
├── @mui/material 9.2.0       (AppLayout)   │
├── @mui/icons-material 9.2.0 (AppLayout)   │
├── @emotion/react 11.14.0    (MUI dep)     │
├── @emotion/styled 11.14.1   (MUI dep)     │
├── axios 1.18.1              (API client)  │
├── react-router-dom 7.18.1   (routing)     │
├── core-js 3.36.0            (polyfill)     │
├── regenerator-runtime 0.14.1 (polyfill)    │
└── es6-promise 4.2.8         (polyfill)     │
                                             │
Office.js (loaded from CDN, not bundled) ───┘
```

---

## File Size Reference

| File | Lines | Est. Tokens | Complexity |
|---|---|---|---|
| `src/taskpane/index.tsx` | 36 | Low | Low |
| `src/taskpane/taskpane.ts` | 18 | Low | Low |
| `src/taskpane/hooks/useEmailContext.ts` | 70 | Medium | Medium |
| `src/taskpane/services/powerflowApi.ts` | 75 | Medium | Medium |
| `src/taskpane/services/mockflow.ts` | 0 | None | None |
| `src/taskpane/components/App.tsx` | 28 | Low | Low |
| `src/taskpane/components/Header.tsx` | 32 | Low | Low |
| `src/taskpane/components/CreateTask/CreateTaskForm.tsx` | 418 | High | High |
| `src/taskpane/components/Layout/AppLayout.tsx` | 134 | Medium | Medium |
| `src/commands/commands.ts` | 35 | Low | Low |
| `webpack.config.js` | 111 | High | Medium |
| `manifest.xml` | 121 | Medium | Medium |
| `package.json` | 82 | Medium | Low |
| `tsconfig.json` | 37 | Low | Low |
| `.vscode/launch.json` | 26 | Low | Low |
| `.vscode/tasks.json` | 160 | Medium | Low |
