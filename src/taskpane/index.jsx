import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require */

const title = "Contoso Task Pane Add-in";

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady((info) => {
  console.log("Office.onReady", info);
  
  // Get taskpaneId from URL if available
  const urlParams = new URLSearchParams(window.location.search);
  const taskpaneId = urlParams.get("taskpaneId");
  console.log("URL taskpaneId:", taskpaneId);
  
  // Check if we were launched from a specific ribbon button
  if (Office.context && Office.context.ui) {
    console.log("Office context available");
  }
  
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
