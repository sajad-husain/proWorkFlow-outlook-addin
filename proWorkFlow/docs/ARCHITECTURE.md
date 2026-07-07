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
│  │  │   │  │  ├─ HeroList       │   │     │       │   │
│  │  │   │  │  └─ TextInsertion  │   │     │       │   │
│  │  │   │  └────────────────────┘   │     │       │   │
│  │  │   └────────────────────────────┘     │       │   │
│  │  └──────────────────────────────────────┘       │   │
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
│  │  ├─ body.setSelectedDataAsync()                  │   │
│  │  └─ notificationMessages.replaceAsync()           │   │
│  └──────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────┘
```

## Project Structure

```
proWorkFlow/
├── src/
│   ├── commands/
│   │   ├── commands.html          # Hidden HTML shell for ribbon command functions
│   │   └── commands.ts            # Ribbon action handler (ExecuteFunction)
│   └── taskpane/
│       ├── index.tsx              # React entry point (bootstraps the app)
│       ├── taskpane.html          # HTML shell loaded in the task pane
│       ├── taskpane.ts            # Office.js integration layer
│       └── components/
│           ├── App.tsx             # Root React component
│           ├── Header.tsx          # Logo + welcome header
│           ├── HeroList.tsx        # Feature list display
│           └── TextInsertion.tsx   # Text input + insert button
├── manifest.xml                    # Add-in manifest (Outlook registration)
├── webpack.config.js               # Build configuration (multi-entry)
├── tsconfig.json                   # TypeScript configuration
├── package.json                    # Dependencies and scripts
├── .vscode/
│   ├── launch.json                 # Debug launch configurations
│   └── tasks.json                  # Build & debug tasks
└── assets/                         # Icons and images
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    ├── icon-80.png
    ├── icon-128.png
    └── logo-filled.png
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
Outlook loads add-in (manifest.xml)
         │
         ▼
Ribbon appears in compose view
         │
         ├── [User clicks "Show Task Pane"]
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
         │     <App title="Contoso Task Pane Add-in" />
         │   </FluentProvider>
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
<FluentProvider theme={webLightTheme}>          [index.tsx:16]
  └── <App title="Contoso Task Pane Add-in">    [App.tsx:19]
        ├── <Header                              [Header.tsx:26]
        │     logo="assets/logo-filled.png"
        │     title={props.title}
        │     message="Welcome" />
        ├── <HeroList                            [HeroList.tsx:44]
        │     message="Discover what this add-in..."
        │     items={[
        │       { icon: <Ribbon24Regular/>, primaryText: "Achieve more..." },
        │       { icon: <LockOpen24Regular/>, primaryText: "Unlock features..." },
        │       { icon: <DesignIdeas24Regular/>, primaryText: "Create and visualize..." }
        │     ]} />
        └── <TextInsertion                       [TextInsertion.tsx:31]
              insertText={insertText} />          [imported from taskpane.ts:3]
```

## Technology Stack

| Layer | Technology | Purpose |
|---|---|---|
| UI Framework | React 18.2 | Component-based UI |
| Design System | Fluent UI React v9 | Theming, components, styling |
| Icons | Fluent UI React Icons v2 | Icon set |
| Language | TypeScript 5.4 (target ES5) | Type safety, IE11 compatibility |
| Build | Webpack 5 | Module bundling, HMR |
| Transpilation | Babel + ts-loader | TypeScript → ES5 |
| Office API | Office.js | Outlook mailbox integration |
| Styling | `makeStyles` (CSS-in-JS) | Component styling via Fluent UI tokens |
| Polyfills | core-js, regenerator-runtime, es6-promise | Legacy browser support |
| Dev Server | webpack-dev-server 5.1 | HTTPS local dev with HMR |
| Dev Certs | office-addin-dev-certs | Auto-generated HTTPS certificates |
| Linting | ESLint (office-addins plugin) | Code quality |
| Debugging | office-addin-debugging | Sideload + launch for testing |

## State Management

Currently uses only **local React state** via `useState`:

| State | Location | Type | Default |
|---|---|---|---|
| `text` (textarea value) | TextInsertion.tsx:32 | `useState<string>` | `"Some text."` |

No global state management is used. The app is a single-view SPA with no routing.

## Manifest Configuration

The add-in is registered with Outlook via `manifest.xml`:

- **Type:** `MailApp` (Outlook add-in)
- **Host:** `Mailbox`
- **Permissions:** `ReadWriteItem`
- **Activation:** Message Read form + Message Compose command surface
- **Minimum Mailbox API:** v1.1 (base), v1.3 (VersionOverrides)

### Ribbon Configuration

```
TabDefault (default compose tab)
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
