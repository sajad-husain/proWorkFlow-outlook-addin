# Code Flow

## Complete Execution Paths

This document traces every code path from entry point to final action.

---

## 1. Task Pane Boot Flow

### Step-by-step trace

```
📄 taskpane.html                          (line 1-41)
   │
   │  HTML page loaded in Outlook task pane webview.
   │  Loads office.js from CDN:
   │    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js">
   │  Provides <div id="container"> for React mount.
   │  Provides <div id="tridentmessage"> fallback for IE/Edge Legacy.
   │
   ▼
📄 index.tsx                               (line 1-30)
   │
   │  Imports:
   │    React, createRoot, App, FluentProvider, webLightTheme
   │
   │  const title = "ProWorkflow for Outlook"         (line 8)
   │  const rootElement = document.getElementById("container")  (line 10)
   │  const root = createRoot(rootElement)              (line 11)
   │
   │  Office.onReady(() => {                            (line 13)
   │    root.render(                                     (line 14)
   │      <FluentProvider theme={webLightTheme}>         (line 15)
   │        <App title={title} />                        (line 16)
   │      </FluentProvider>
   │    );
   │  });
   │
   │  HMR support:                                       (line 21-30)
   │    if (module.hot) { hot.accept("./components/App", ...) }
   │
   ▼
📄 App.tsx                                 (line 1-32)
   │
   │  Imports:
   │    Header, CreateTaskForm, useEmailContext, makeStyles
   │
   │  const { emailData, loading } = useEmailContext()  (line 22)
   │    ↓
   │  Hook fires on mount — extracts email data from the
   │  current Outlook item (subject, body, from, attachments)
   │    ↓
   │  Renders:                                           (line 24-28)
   │    <div className={styles.root}>
   │      <Header logo="assets/logo-filled.png"
   │              title={props.title} />
   │      <CreateTaskForm emailData={emailData}
   │                      loading={loading} />
   │    </div>
   │
   ▼
┌─────────────────────────────────────────────────────────────┐
│  Component Tree                                              │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  📄 Header.tsx                                                │
│    Props: { title, logo }                                     │
│    Renders: <Image> + <h1> with title                         │
│                                                               │
│  📄 CreateTaskForm.tsx (line 125-416)                        │
│    Props: { emailData, loading }                              │
│    States: formData, isSubmitting                             │
│    Renders:                                                   │
│      ├─ <Toaster /> (notification toasts)                    │
│      ├─ Email preview (subject, from)                        │
│      ├─ Workspace <Select>                                   │
│      ├─ Task list <Select> (filtered by workspace)           │
│      ├─ Task <Select> (for editing)                          │
│      ├─ Title <Input> (pre-filled from email subject)        │
│      ├─ Description <Textarea> (pre-filled from email body)  │
│      ├─ Assignee <Select> + Due date <Input type="date">    │
│      ├─ Urgent checkbox + Add attachments checkbox           │
│      └─ Submit <button>                                      │
│                                                               │
└─────────────────────────────────────────────────────────────┘
```

### State Initialization

```
hooks/useEmailContext.ts:12
  const [emailData, setEmailData] = useState<EmailData | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  On mount (useEffect):
    → Reads item.subject (sync)
    → Reads item.from.emailAddress (sync)
    → Calls item.body.getAsync("text") (async)       [line 30-35]
    → Reads item.attachments (sync)
    → setEmailData({ subject, body, from, attachments })
    → setLoading(false)
    → Fallback mock data if Office.context.mailbox is unavailable [line 58-65]

CreateTaskForm.tsx:148
  const [formData, setFormData] = useState({
    workspace: "", taskList: "", task: "", title: "",
    description: "", assignee: "", dueDate: "",
    isUrgent: false, addAttachments: false
  });

  useEffect pre-fills title & description from emailData [line 163-171]:
    When emailData changes → setFormData({ ...prev, title: emailData.subject, description: emailData.body })
```

---

## 2. Task Creation Flow (User Action)

```
User opens task pane → email data loads automatically
  → title & description pre-filled from email subject/body
  → User selects workspace, task list, assignee, due date
  → User optionally checks "Mark as Urgent" / "Add attachments"

User clicks "Create task" button
  → handleSubmit(e)                          [CreateTaskForm.tsx:191]
  → setIsSubmitting(true)
  → console.log("Creating task:", formData)   [line 197]
  → Simulates API call (1.5s delay)           [line 200]
      ↓
    (FUTURE: calls proworkflowApi.createTask())
    📄 services/proworkflowApi.ts
      → if USE_MOCK: returns mockTaskCreationResponse  [line 57-61]
      → if !USE_MOCK: axios.post('/tasks', taskData)   [line 63-67]
      ↓
  → On success: dispatchToast("Task Created Successfully!") [line 202-208]
  → On error: dispatchToast("Error Creating Task")        [line 222-228]
  → setFormData resets to defaults                         [line 211-220]
  → setIsSubmitting(false)                                 [line 230]
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
┌──────────────┐
│  Outlook      │  Outlook opens add-in on message read
│  Mailbox API  │
└──────┬───────┘
       │ Office.context.mailbox.item
       ▼
┌─────────────────────────────────────┐
│  useEmailContext hook                │
│  ├─ item.subject (sync)             │
│  ├─ item.from.emailAddress (sync)   │
│  ├─ item.body.getAsync("text")      │
│  ├─ item.attachments (sync)         │
│  └─ returns { emailData, loading }  │
└──────────────┬──────────────────────┘
               │ emailData: { subject, body, from, ... }
               ▼
┌─────────────────────────────────────┐
│  CreateTaskForm                      │
│  ├─ useEffect pre-fills title/desc  │
│  ├─ User fills remaining fields     │
│  └─ handleSubmit() on click         │
└──────────────┬──────────────────────┘
               │ formData
               ▼
┌─────────────────────────────────────┐          ┌──────────────────────┐
│  proworkflowApi.createTask()        │──if────▶│  Mock response       │
│  (USE_MOCK = true by default)       │  MOCK    │  (1.5s delay)        │
│                                     │─────────│                      │
│  axios.post('/tasks', formData)     │  REAL    │  ProWorkflow API     │
│                                     │          │  (api.proworkflow.net)│
└──────────────┬──────────────────────┘          └──────────────────────┘
               │ result
               ▼
┌─────────────────────────────────────┐
│  Toast notification (success/error) │
│  └─ Form reset on success          │
└─────────────────────────────────────┘
```

---

## 7. Import/Export Graph

```
index.tsx
  ├── import App from "./components/App"
  ├── import { FluentProvider, webLightTheme } from "@fluentui/react-components"
  ├── import { createRoot } from "react-dom/client"

App.tsx
  ├── import Header from "./Header"
  ├── import CreateTaskForm from "./CreateTask/CreateTaskForm"
  ├── import { useEmailContext } from "../hooks/useEmailContext"
  └── import { makeStyles } from "@fluentui/react-components"

CreateTaskForm.tsx
  ├── import { useState, useEffect } from "react"
  ├── import { makeStyles, tokens, Text, Input, Textarea, Button,
  │          Select, Checkbox, Card, Spinner, Field, Divider,
  │          Toast, ToastTitle, ToastBody, Toaster,
  │          useToastController } from "@fluentui/react-components"
  └── import { Add24Regular, Calendar24Regular, Person24Regular,
              Flag24Regular, CheckmarkCircle24Regular } from "@fluentui/react-icons"

Header.tsx
  └── import { Image, tokens, makeStyles } from "@fluentui/react-components"

hooks/useEmailContext.ts
  ├── import { useState, useEffect } from "react"
  └── (uses global Office — no import)

services/proworkflowApi.ts
  ├── import axios from "axios"
  └── import { mockProjects, mockAssignees, mockTaskCreationResponse } from "./mockData"

AppLayout.tsx (standalone, not yet used in main flow)
  ├── import { AppBar, Toolbar, Typography, Container, Box,
  │          IconButton, Drawer, List, ListItem, ListItemIcon,
  │          ListItemText, Divider } from "@mui/material"
  ├── import MenuIcon from "@mui/icons-material/Menu"
  ├── import TaskIcon from "@mui/icons-material/Assignment"
  ├── import EditIcon from "@mui/icons-material/Edit"
  ├── import CloseIcon from "@mui/icons-material/Close"
  └── import { useNavigate } from "react-router-dom"

taskpane.ts  ←── shared Office.js utility
  └── (no imports)

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
       ├── useEmailContext fails (hooks/useEmailContext.ts:49-51)
       │     → catch(error) { console.error('Error loading email:', error) }
       │     → emailData remains null, loading → false
       │     → Form renders with empty fields (no pre-fill)
       │
       ├── form submit fails (CreateTaskForm.tsx:221-228)
       │     → catch(error)
       │     → dispatchToast("Error Creating Task", { intent: "error" })
       │     → User sees error toast notification
       │
       ├── proworkflowApi fails (services/proworkflowApi.ts)
       │     → catch(error) { console.error('API Error:', error); throw error; }
       │     → Error propagates to submit handler
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

The following are planned enhancements:

```
1. Connect real ProWorkflow API:
   → Set USE_MOCK = false in services/proworkflowApi.ts
   → Add real API credentials
   → Wire up axios calls instead of mock data

2. Multi-view routing:
   → AppLayout.tsx with React Router DOM
   → Views: Create Task, Edit Task, View Tasks, Settings
   → Drawer navigation already built in AppLayout

3. Attachment support:
   → Implement "Add attachments" checkbox logic
   → Upload email attachments to ProWorkflow tasks

4. Error boundaries:
   → Add React error boundary component
   → Graceful fallback UI instead of white screen

5. Authentication:
   → ProWorkflow OAuth or API key management UI
   → Persistent auth state
```
