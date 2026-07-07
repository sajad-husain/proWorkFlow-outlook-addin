import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./components/App";

/* global document, Office, module, require */

const title = "ProWorkflow for Outlook";

const rootElement = document.getElementById("container");
if (!rootElement) {
  throw new Error("Root element not found");
}
const root = createRoot(rootElement);
declare const module: {
  hot?: {
    accept(path: string, callback: () => void): void;
  };
};

Office.onReady(() => {
  root.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root.render(
      <FluentProvider theme={webLightTheme}>
        <NextApp title={title} />
      </FluentProvider>
    );
  });
}
