# Code Flow

## Complete Execution Paths

This document traces every code path from entry point to final action.

---

## 1. Task Pane Boot Flow

### Step-by-step trace

```
📄 taskpane.html                          (line 1-26)
   │
   │  HTML page loaded in Outlook task pane webview.
   │  Loads office.js from CDN:
   │    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js">
   │  Provides <div id="container"> for React mount.
   │  Provides <div id="tridentmessage"> fallback for IE/Edge Legacy.
   │
   ▼
📄 index.tsx                               (line 1-27)
   │
   │  Imports:
   │    React, createRoot, App, FluentProvider, webLightTheme
   │
   │  const title = "Contoso Task Pane Add-in"        (line 8)
   │  const rootElement = document.getElementById("container")  (line 10)
   │  const root = createRoot(rootElement)              (line 11)
   │
   │  Office.onReady(() => {                            (line 14)
   │    // Waits for Office.js to fully initialize.
   │    // The callback fires once Office is ready.
   │    root.render(                                     (line 15)
   │      <FluentProvider theme={webLightTheme}>         (line 16)
   │        <App title={title} />                        (line 17)
   │      </FluentProvider>
   │    );
   │  });
   │
   │  HMR support:                                       (line 22-27)
   │    if (module.hot) { hot.accept("./components/App", ...) }
   │
   ▼
📄 App.tsx                                 (line 1-47)
   │
   │  Imports:
   │    Header, HeroList, HeroListItem, TextInsertion
   │    makeStyles, Fluent UI icons (Ribbon24Regular, LockOpen24Regular,
   │    DesignIdeas24Regular)
   │    insertText from "../taskpane"
   │
   │  const listItems: HeroListItem[] = [                (line 23-36)
   │    { icon: <Ribbon24Regular />, primaryText: "Achieve more..." },
   │    { icon: <LockOpen24Regular />, primaryText: "Unlock features..." },
   │    { icon: <DesignIdeas24Regular />, primaryText: "Create and visualize..." }
   │  ]
   │
   │  Renders:                                           (line 38-44)
   │    <div className={styles.root}>                    (line 39)
   │      <Header logo="assets/logo-filled.png"          (line 40)
   │              title="Contoso Task Pane Add-in"
   │              message="Welcome" />
   │      <HeroList message="Discover what..."           (line 41)
   │               items={listItems} />
   │      <TextInsertion insertText={insertText} />      (line 42)
   │    </div>
   │
   ▼
┌─────────────────────────────────────────────────────────────┐
│  Component Tree (rendered concurrently — no nesting depth)   │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  📄 Header.tsx  (line 26-36)                                  │
│    Props: { title, logo, message }                             │
│    Renders:                                                    │
│      <section className={styles.welcome__header}>             │
│        <Image width="90" height="90" src={logo} alt={title} />│
│        <h1 className={styles.message}>{message}</h1>          │
│      </section>                                                │
│                                                               │
│  📄 HeroList.tsx  (line 44-62)                                │
│    Props: { message: string, items: HeroListItem[] }           │
│    Maps items → <li> with icon + primaryText                  │
│    Renders:                                                    │
│      <div className={styles.welcome__main}>                  │
│        <h2>{message}</h2>                                     │
│        <ul>                                                    │
│          {items.map((item, index) => (                         │
│            <li key={index}>                                    │
│              <i>{item.icon}</i>                                │
│              <span>{item.primaryText}</span>                   │
│            </li>                                               │
│          ))}                                                   │
│        </ul>                                                   │
│      </div>                                                    │
│                                                               │
│  📄 TextInsertion.tsx  (line 31-54)                           │
│    Props: { insertText: (text: string) => void }               │
│    State: text = useState("Some text.")                        │
│    Renders:                                                    │
│      <div className={styles.textPromptAndInsertion}>          │
│        <Field label="Enter text...">                           │
│          <Textarea value={text} onChange={handleTextChange} /> │
│        </Field>                                                │
│        <Field>Click the button to insert text.</Field>         │
│        <Button onClick={handleTextInsertion}>Insert text</Button>│
│      </div>                                                    │
│                                                               │
└─────────────────────────────────────────────────────────────┘
```

### State Initialization

```
TextInsertion.tsx:32
  const [text, setText] = useState<string>("Some text.");

  When user types in the Textarea:
    handleTextChange(event)          (line 38-40)
      → setText(event.target.value)  (line 39)
      → React re-renders Textarea with new value
```

---

## 2. Text Insertion Flow (User Action)

```
User types in Textarea
  → handleTextChange(event)                     [TextInsertion.tsx:38]
  → setText(event.target.value)                 [TextInsertion.tsx:39]
  → Component re-renders with updated text

User clicks "Insert text" Button
  → handleTextInsertion()                       [TextInsertion.tsx:34]
  → await props.insertText(text)                [TextInsertion.tsx:35]
      ↓
    📄 taskpane.ts  (line 3-17)
      insertText(text)
      → Office.context.mailbox.item             [line 6]
          ?.body.setSelectedDataAsync(           [line 6-13]
            text,
            { coercionType: Office.CoercionType.Text },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                throw asyncResult.error.message;
              }
            }
          )
      → Text inserted at cursor in email body
      → On failure: error logged to console     [line 16]
```

---

## 3. Command Function Flow (Ribbon Button)

```
User clicks "Perform an action" ribbon button
  → Outlook loads commands.html (hidden iframe)  [manifest.xml:45]
  → Office.js initializes
  → commands.ts fires:
      Office.onReady(() => {})                   [line 8]
      // No-op: just ensures Office is ready
  → Outlook invokes the registered function:
      action(event)                              [line 16-32]
      ↓
    Creates notification message:                [line 17-22]
      type: "InformationalMessage"
      message: "Performed action."
      icon: "Icon.80x80"
      persistent: true
      ↓
    Office.context.mailbox.item                  [line 25]
      ?.notificationMessages.replaceAsync(       [line 25-28]
        "ActionPerformanceNotification",
        message
      )
      → Shows notification in email compose form
      ↓
    event.completed()                            [line 31]
      → Tells Outlook the command is done
      ↓
  Function was registered at import time:        [line 35]
    Office.actions.associate("action", action)
```

---

## 4. Webpack Build Flow

```
npm run build (or build:dev / watch / dev-server)
  │
  ▼
📄 webpack.config.js  (line 16-110)
  │
  │  mode = "production" | "development"
  │  devtool = "source-map"
  │
  │  Entry points:                              [line 20-28]
  │    polyfill:  core-js + regenerator-runtime
  │    react:     react + react-dom
  │    taskpane:  index.tsx + taskpane.html (depends on react)
  │    commands:  commands.ts
  │
  │  Module Rules:                              [line 36-61]
  │    .ts  → babel-loader (ES5 transpilation)
  │    .tsx → ts-loader    (TypeScript compilation)
  │    .html → html-loader (inline HTML as strings)
  │    .png/.jpg/etc → asset/resource (copies to assets/)
  │
  │  Plugins:                                   [line 63-95]
  │    HtmlWebpackPlugin(taskpane.html)
  │      → injects polyfill + taskpane + react chunks
  │    HtmlWebpackPlugin(commands.html)
  │      → injects polyfill + commands chunks
  │    CopyWebpackPlugin
  │      → copies assets/* and manifest*.xml to output
  │      → in production: replaces localhost URLs with urlProd
  │    ProvidePlugin
  │      → provides Promise polyfill (es6-promise)
  │
  │  DevServer:                                 [line 97-107]
  │    port: 3000
  │    https: auto-generated dev certs
  │    hot: true (HMR)
  │    CORS: Access-Control-Allow-Origin: *
  │
  ▼
📁 dist/
  ├── taskpane.html        (injected with polyfill, taskpane, react <script> tags)
  ├── commands.html        (injected with polyfill, commands <script> tags)
  ├── polyfill.js
  ├── react.js
  ├── taskpane.js
  ├── commands.js
  ├── assets/              (icons, images)
  └── manifest.xml         (copied, URLs replaced in production)
```

### Modes

| Mode | Command | Source Maps | HMR | Minification |
|---|---|---|---|---|
| Development | `npm run build:dev` | Full | No | No |
| Production | `npm run build` | Full | No | Yes |
| Dev Server (HMR) | `npm run dev-server` | Full | Yes | No |
| Watch | `npm run watch` | Full | No | No |

---

## 5. Debugging Launch Flow

```
VS Code: "Outlook Desktop (Edge Chromium)" launch config
  │
  ▼
📄 .vscode/tasks.json
  preLaunchTask: "Debug: Outlook Desktop"
    │
    ▼
  npm run start -- desktop --app outlook
    │
    ▼
📄 package.json: scripts.start
  office-addin-debugging start manifest.xml
    │
    ├── 1. Builds the project (webpack --mode development)
    ├── 2. Starts dev server (webpack serve --mode development)
    ├── 3. Registers the add-in in Outlook (sideloading)
    ├── 4. Launches Outlook with the add-in loaded
    └── 5. Opens debugging port (9229 for Edge Chromium)
    │
    ▼
📄 .vscode/launch.json
  type: msedge
  request: attach
  port: 9229
  webRoot: ${workspaceRoot}
    │
    ▼
  VS Code attaches debugger to the Outlook webview
  → Breakpoints in .ts/.tsx files are hit via source maps
  → Console output appears in VS Code debug console
```

---

## 6. Data Flow Diagram

```
┌──────────────┐     Typing     ┌──────────────────┐
│   Textarea    │ ──────────────▶│  handleTextChange │
│  (controlled) │◀──────────────│  setText(value)   │
└──────────────┘  re-render     └──────────────────┘
       │
       │ Click "Insert text"
       ▼
┌──────────────────┐
│ handleTextInsert │
│   await insert   │
│   Text(text)     │
└──────┬───────────┘
       │
       ▼
┌─────────────────────────────────────┐
│  taskpane.ts: insertText(text)      │
│  Office.context.mailbox.item.body   │
│  .setSelectedDataAsync(text, ...)   │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Office.js API Bridge               │
│  → Outlook writes text to email     │
│  → AsyncResult returned to callback │
└─────────────────────────────────────┘
```

---

## 7. Import/Export Graph

```
index.tsx
  ├── import App from "./components/App"
  ├── import { FluentProvider, webLightTheme } from "@fluentui/react-components"
  ├── import { createRoot } from "react-dom/client"
  │
App.tsx
  ├── import Header from "./Header"
  ├── import HeroList, { HeroListItem } from "./HeroList"
  ├── import TextInsertion from "./TextInsertion"
  ├── import { insertText } from "../taskpane"
  ├── import { makeStyles } from "@fluentui/react-components"
  └── import { icons } from "@fluentui/react-icons"
      │
Header.tsx
  ├── import { Image, tokens, makeStyles } from "@fluentui/react-components"
  │
HeroList.tsx
  ├── import { tokens, makeStyles } from "@fluentui/react-components"
  │
TextInsertion.tsx
  ├── import { useState } from "react"
  └── import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components"
      │
taskpane.ts  ←── shared Office.js logic
  └── (no imports)
      │
commands.ts  ←── separate bundle, no dependency on taskpane
  └── (no imports)
```

## 8. Error Handling Flow

```
┌─────────────┐
│  Task Pane   │
│  Operations  │
└──────┬──────┘
       │
       ├── Office.onReady() fails
       │     → Nothing rendered (root.render never called)
       │     → Check browser console for Office.js errors
       │
       ├── insertText() fails (taskpane.ts:10-12)
       │     → asyncResult.status === Office.AsyncResultStatus.Failed
       │     → throw asyncResult.error.message
       │     → catch: console.log("Error: " + error)
       │     → User sees no feedback (silent failure)
       │
       └── Component render error
             → React error boundary needed (not yet implemented)
             → "Something went wrong" white screen in task pane
       │
┌─────────────┐
│  Commands    │
│  Operations  │
└──────┬──────┘
       │
       ├── notificationMessages.replaceAsync() fails
       │     → No error handling (callback omitted)
       │     → Action completed regardless
       │
       └── event.completed() not called
             → Outlook shows "Add-in is taking too long" warning
             → Happens if action() throws before event.completed()

```

## 9. Future Growth Paths

The following patterns are expected to be added for ProWorkflow integration:

```
User views email in Outlook
  → Add-in extracts email details (sender, subject, body)
  → Data formatted as ProWorkflow task payload
  → Axios POST to ProWorkflow API
  → Response displayed in task pane

Expected new data flow:
  App.tsx
    → gets email context from Office.context.mailbox.item
    → calls a new api/ service module
    → api module uses axios to call ProWorkflow
    → response flows back through React state
    → UI updates with task creation result
```
