import * as React from "react";
import { createRoot } from "react-dom/client";
import DialogApp from "./DialogApp";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const title = "Contoso Dialog Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  // Listen for localStorage messages before rendering the dialog
  console.log("[dialog.tsx] Office is ready, setting up storage event listener...");
  function handleStorageEvent(event: StorageEvent) {
    if (event.key === "dialog_command_message") {
      try {
        const messageData = event.newValue ? JSON.parse(event.newValue) : null;
        if (messageData) {
          // Log the message and show a notification (if needed)
          console.log("[dialog.tsx] Storage event received:", messageData);
          // You can add Office notification logic here if needed
        }
        // Clear the item to avoid re-processing
        localStorage.removeItem("dialog_command_message");
      } catch (error) {
        console.error("[dialog.tsx] Could not parse message from localStorage: ", error);
      }
    }
  }
  window.addEventListener("storage", handleStorageEvent);

  root?.render(
    <FluentProvider theme={webLightTheme}>
      <DialogApp />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./DialogApp", () => {
    const NextDialogApp = require("./DialogApp").default;
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <NextDialogApp />
      </FluentProvider>
    );
  });
}
