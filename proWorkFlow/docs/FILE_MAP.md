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
| **Lines** | 27 |
| **Key line** | `14: Office.onReady(...)` — app boot after Office.js init |
| **HMR** | Lines 22-27: `module.hot.accept` |

---

### `src/taskpane/taskpane.ts`
**Purpose:** Office.js integration layer. Contains the function that writes text to the email compose surface.

| Detail | Value |
|---|---|
| **Imports** | None (uses global `Office` and `console`) |
| **Exports** | `insertText(text: string): Promise<void>` |
| **Lines** | 18 |
| **Key line** | `6: Office.context.mailbox.item?.body.setSelectedDataAsync(...)` |
| **Error handling** | Lines 10-12: throws on `AsyncResultStatus.Failed` |
| **Used by** | `App.tsx` (line 7 imports and passes to `TextInsertion`) |

---

### `src/taskpane/components/App.tsx`
**Purpose:** Root React component. Assembles the UI from child components.

| Detail | Value |
|---|---|
| **Imports** | `Header`, `HeroList` (+ `HeroListItem`), `TextInsertion`, `@fluentui/react-components` (`makeStyles`), `@fluentui/react-icons` (`Ribbon24Regular`, `LockOpen24Regular`, `DesignIdeas24Regular`), `insertText` from `../taskpane` |
| **Exports** | `App(props: AppProps)` |
| **Props** | `{ title: string }` |
| **Lines** | 47 |
| **Key line** | `42: <TextInsertion insertText={insertText} />` — passes Office.js function to UI |
| **Static data** | Lines 23-36: `HeroListItem[]` with Fluent UI icons |

---

### `src/taskpane/components/Header.tsx`
**Purpose:** Displays the add-in logo and welcome message.

| Detail | Value |
|---|---|
| **Imports** | `@fluentui/react-components` (`Image`, `tokens`, `makeStyles`) |
| **Exports** | `Header(props: HeaderProps)` |
| **Props** | `{ title: string, logo: string, message: string }` |
| **Lines** | 38 |
| **Style** | Fluent UI tokens-based CSS-in-JS via `makeStyles` |

---

### `src/taskpane/components/HeroList.tsx`
**Purpose:** Displays a list of feature highlights with icons.

| Detail | Value |
|---|---|
| **Imports** | `@fluentui/react-components` (`tokens`, `makeStyles`) |
| **Exports** | `HeroList` (default), `HeroListItem` (interface) |
| **Props** | `{ message: string, items: HeroListItem[] }` |
| **HeroListItem** | `{ icon: React.JSX.Element, primaryText: string }` |
| **Lines** | 62 |
| **Key line** | `48: items.map(...)` — renders the list items |

---

### `src/taskpane/components/TextInsertion.tsx`
**Purpose:** Provides the text input UI and insert button.

| Detail | Value |
|---|---|
| **Imports** | `react` (`useState`), `@fluentui/react-components` (`Button`, `Field`, `Textarea`, `tokens`, `makeStyles`) |
| **Exports** | `TextInsertion(props: TextInsertionProps)` |
| **Props** | `{ insertText: (text: string) => void }` |
| **State** | `useState<string>("Some text.")` — controlled textarea value |
| **Lines** | 57 |
| **Key line** | `34: handleTextInsertion` — calls `props.insertText(text)` on button click |
| **Key line** | `38: handleTextChange` — updates state on textarea input |

---

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
| **Id** | `22e64ad0-fb04-48d5-bf5f-eb99d21d1be9` |
| **DisplayName** | `proWorkFlow` |
| **Permissions** | `ReadWriteItem` |
| **Requirement** | Mailbox API v1.1 (base), v1.3 (VersionOverrides) |
| **Tasks** | `ShowTaskpane` → taskpane.html, `ExecuteFunction` → `action()` |
| **Lines** | 108 |

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
| **Dependencies** | React 18, Fluent UI v9, MUI v9, Axios, React Router DOM v7 |
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

---

## Dependency Graph (npm packages)

```
react 18.2 ─────────────────────────────────┐
react-dom 18.2 ─────────────────────────────┤
@fluentui/react-components 9.55.1 ──────────┤
@fluentui/react-icons 2.0.264 ──────────────┤
├── @mui/material 9.2.0       (unused)      │
├── @mui/icons-material 9.2.0 (unused)      │
├── axios 1.18.1              (unused)       │
├── react-router-dom 7.18.1   (unused)       │
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
| `src/taskpane/index.tsx` | 27 | Low | Low |
| `src/taskpane/taskpane.ts` | 18 | Low | Low |
| `src/taskpane/components/App.tsx` | 47 | Medium | Low |
| `src/taskpane/components/Header.tsx` | 38 | Low | Low |
| `src/taskpane/components/HeroList.tsx` | 62 | Medium | Low |
| `src/taskpane/components/TextInsertion.tsx` | 57 | Medium | Medium |
| `src/commands/commands.ts` | 35 | Low | Low |
| `webpack.config.js` | 111 | High | Medium |
| `manifest.xml` | 108 | Medium | Medium |
| `package.json` | 82 | Medium | Low |
| `tsconfig.json` | 37 | Low | Low |
| `.vscode/launch.json` | 26 | Low | Low |
| `.vscode/tasks.json` | 160 | Medium | Low |
