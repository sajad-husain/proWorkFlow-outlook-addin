# Debugging Guide

## Overview

Debugging an Outlook add-in requires three concurrent processes:
1. **Webpack Dev Server** - serves the add-in files via HTTPS on localhost:3000
2. **Outlook Desktop** - hosts the add-in in its webview
3. **VS Code Debugger** - attaches to the webview for breakpoints/source maps

---

## Quick Start

### Prerequisites

- Node.js 18+
- Outlook Desktop (Win32)
- VS Code with recommended extensions:
  - `ms-edgedevtools.vscode-edge-devtools`
  - `msoffice.microsoft-office-add-in-debugger`
  - `dbaeumer.vscode-eslint`

### First-Time Setup

```powershell
# In the proWorkFlow/ directory:
npm install               # Install dependencies
npm run signin            # Sign in to M365 (required for sideloading)
```

### Start Debugging

```powershell
# Option 1: VS Code (recommended)
#   1. Open proWorkFlow/ in VS Code
#   2. Run Tasks → "Dev Server" (starts webpack dev server)
#   3. Press F5 → "Outlook Desktop (Edge Chromium)"
#
# Option 2: Command line
npm start                 # Builds, starts dev server, launches Outlook
```

---

## Debugging Configurations

### VS Code Launch Configurations

Two debug profiles are defined in `.vscode/launch.json`:

#### 1. Outlook Desktop (Edge Chromium) — Primary

| Field | Value |
|---|---|
| `type` | `msedge` |
| `request` | `attach` |
| `port` | `9229` |
| `preLaunchTask` | `Debug: Outlook Desktop` |
| `postDebugTask` | `Stop Debug` |
| `webRoot` | `${workspaceRoot}` |

This is the modern configuration targeting Outlook's Edge Chromium webview (Outlook 365 v2303+).

#### 2. Outlook Desktop (Edge Legacy) — Fallback

| Field | Value |
|---|---|
| `type` | `office-addin` |
| `request` | `attach` |
| `port` | `9222` |
| `url` | `https://localhost:3000/taskpane.html?_host_Info=Outlook$Win32$16.01$en-US$$$$0` |

Use this for older Outlook builds using EdgeHTML/Edge Legacy webview.

---

## Debugging Flow Diagram

```
VS Code F5
  │
  ▼
preLaunchTask: "Debug: Outlook Desktop"
  │
  ▼
npm run start -- desktop --app outlook
  │
  ├── 1. webpack --mode development          [Builds project with source maps]
  ├── 2. webpack serve --mode development    [Starts HTTPS dev server :3000]
  ├── 3. office-addin-debugging              [Sideloads add-in into Outlook]
  ├── 4. Launch Outlook                      [Opens Outlook with add-in]
  └── 5. Open debugging port                 [Port 9229 for Edge Chromium]
  │
  ▼
VS Code attaches to webview
  │
  ▼
Breakpoints in .ts/.tsx files are hit
  │
  ▼
postDebugTask: "Stop Debug"
  │
  ▼
npm run stop                                 [Removes sideloaded add-in]
```

---

## Setting Breakpoints

### TypeScript Source Files

Because `sourceMap: true` is set in `tsconfig.json` and `devtool: "source-map"` is set in `webpack.config.js`, breakpoints set directly in `.ts`/`.tsx` files will map correctly to the running code.

**Files where breakpoints are useful:**

| File | Best Breakpoint Location | What to Inspect |
|---|---|---|
| `index.tsx:13` | `Office.onReady(() => {` | Is Office.js initializing? |
| `hooks/useEmailContext.ts:19` | `const item = Office.context.mailbox.item` | Is the Outlook item accessible? |
| `hooks/useEmailContext.ts:30` | `item.body?.getAsync('text', ...)` | Is the email body being extracted? |
| `services/proworkflowApi.ts:28` | `getProjects` | Is the API being called? |
| `services/proworkflowApi.ts:56` | `createTask` | Is task creation triggered? |
| `commands.ts:16` | `function action(event)` | Is the command function triggered? |
| `commands.ts:25` | `Office.context.mailbox.item?.notificationMessages.replaceAsync(...)` | Is the notification being set? |
| `CreateTaskForm.tsx:163` | `useEffect` pre-fill | Is the email data being used to pre-fill? |
| `CreateTaskForm.tsx:191` | `handleSubmit` | Is the form submission handled? |
| `App.tsx:22` | `const { emailData, loading } = useEmailContext()` | Are the components rendering correctly? |

### CSS-in-JS Breakpoints

Fluent UI `makeStyles` generates dynamic CSS at runtime. To inspect styles:
- Use the browser DevTools Elements panel (opened via VS Code Edge DevTools)
- Look for injected `<style>` tags with generated class names

---

## Debugging Console

### What to Watch

| Console Source | What Appears | How to View |
|---|---|---|
| `webpack dev server` | Build errors, HMR status | VS Code terminal (Dev Server task) |
| `Outlook webview console` | `console.log()` from add-in code | VS Code Debug Console (attached debugger) |
| `Office.js errors` | Initialization failures, API errors | VS Code Debug Console |

### Enabling Verbose Office.js Logging

Office.js has internal logging that can be enabled:

```typescript
// In index.tsx or commands.ts, before Office.onReady():
Office.onReady((info) => {
  console.log("Office.js ready. Host:", info.host, "Platform:", info.platform);
  // ...
});
```

---

## Common Issues & Fixes

### 1. White Screen / App Not Loading

```
Symptoms:  Task pane opens but shows blank white screen
Root Cause: One of:
  - Office.js CDN not loading (no internet / blocked)
  - React render error (uncaught exception)
  - Polyfill issue in legacy webview

Diagnosis:
  1. Open DevTools (VS Code Edge DevTools)
  2. Check Console tab for errors
  3. Check Network tab for office.js CDN load

Fix:
  - Check internet connectivity
  - Add React error boundary
  - Verify polyfill entries in webpack.config.js
```

### 2. Trident Message Appears

```
Symptoms:  "This add-in is not supported in this version of Office"
Root Cause: Outlook is using IE/Edge Legacy webview

Diagnosis:
  - Check #tridentmessage visibility in taskpane.html

Fix:
  - Update Outlook to use Edge Chromium webview
  - Office 365 → File → Options → General → "Enable Microsoft Edge WebView2"
  - Or deploy to Outlook on Mac/Web which use modern webviews
```

### 3. "Add-in Not Found" / Sideload Fails

```
Symptoms:  Outlook doesn't show the add-in after npm start
Root Cause: One of:
  - M365 account not signed in
  - Manifest validation failure
  - Dev server not running
  - HTTPS certificate issue

Diagnosis:
  npm run validate              # Check manifest
  npm run signin                # Verify auth

Fix:
  npm run signin                # Sign in to M365
  npm run validate              # Fix any manifest errors
  Check dev server is running   # localhost:3000 should respond
  Check certs:                  # office-addin-dev-certs install
```

### 4. Source Maps Not Working / Breakpoints Not Hit

```
Symptoms:  Breakpoints show as "unbound" or grayed out
Root Cause: One of:
  - Debugger attached before webview loaded
  - Wrong debug port
  - Source map not generated

Diagnosis:
  - Check dist/ for .map files
  - Verify tsconfig.json has sourceMap: true
  - Verify webpack.config.js has devtool: "source-map"

Fix:
  - Restart debug session (F5 again)
  - Check port 9229 is listening
  - Rebuild with source maps
```

### 5. HMR Not Working

```
Symptoms:  Changes don't reflect without full page reload
Root Cause: One of:
  - Outdated webpack config
  - Module.hot not accepted

Diagnosis:
  - Check browser console for HMR status messages

Fix:
  - Ensure hot.accept is called in index.tsx:22-27
  - Verify webpack.config.js has hot: true
  - Restart dev server
```

### 6. "Perform an action" Doesn't Show Notification

```
Symptoms:  Ribbon button clicked but no notification appears
Root Cause: One of:
  - commands.html not loaded
  - action() function not registered
  - event.completed() called before async operation finishes

Diagnosis:
  - Set breakpoint in commands.ts:16
  - Check Console for errors

Fix:
  - Verify Office.actions.associate("action", action) at commands.ts:35
  - Ensure event.completed() is called after async operations
```

### 7. port 3000 Already in Use

```
Symptoms:  Dev server fails to start
Root Cause: Another process using port 3000

Fix:
  # Option 1: Kill the process
  netstat -ano | findstr :3000
  taskkill /PID <pid> /F

  # Option 2: Use a different port
  # Edit package.json config.dev_server_port
  # Or set env var: $env:npm_package_config_dev_server_port=3001
```

---

## Available Scripts Reference

| Script | Purpose | When to Use |
|---|---|---|
| `npm start` | Build + launch + sideload | Full debug session |
| `npm run dev-server` | Start dev server only | When you want to keep the server running separately |
| `npm stop` | Stop debugging + remove sideload | End debug session |
| `npm run build:dev` | Build for development (with source maps) | Check for build errors |
| `npm run build` | Build for production (minified) | Release |
| `npm run watch` | Rebuild on file changes | Editing without dev server |
| `npm run validate` | Validate manifest.xml | Check manifest correctness |
| `npm run lint` | Run ESLint check | Before committing |
| `npm run lint:fix` | Auto-fix lint issues | Quick formatting fixes |
| `npm run signin` | Sign in to M365 | First-time setup |
| `npm run signout` | Sign out of M365 | Clean up |

---

## VS Code Tasks Reference

Available from Terminal → Run Task:

| Task Name | Action |
|---|---|
| `Build (Development)` | `npm run build:dev` |
| `Build (Production)` | `npm run build` |
| `Dev Server` | `npm run dev-server` |
| `Watch` | `npm run watch` |
| `Debug: Outlook Desktop` | Pre-launch for F5 debugging |
| `Debug: Excel Desktop` | (context: for other Office apps) |
| `Lint: Check for problems` | `npm run lint` |
| `Lint: Fix all auto-fixable problems` | `npm run lint:fix` |
| `Stop Debug` | `npm run stop` |
| `Install` | `npm install` |

---

## Advanced: Manual Debugging

If VS Code debugging doesn't work, you can debug manually:

```powershell
# 1. Start dev server
npm run dev-server

# 2. In a separate terminal, start Outlook with the add-in
npm start

# 3. Open Edge DevTools and attach to Outlook's webview process:
#    - Open edge://inspect in Microsoft Edge
#    - Find the Outlook webview process
#    - Click "inspect"

# 4. When done:
npm stop
```

## Advanced: Production Debugging

To debug a deployed add-in:
1. Open the add-in in Outlook
2. Open DevTools (right-click in task pane → Inspect, or Ctrl+Shift+I)
3. Source maps must be deployed alongside the build (they are generated by default)
4. Debug as you would a normal web application
