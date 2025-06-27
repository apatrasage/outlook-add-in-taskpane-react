import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const title = "Contoso Task Pane Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  // Listen for localStorage messages before rendering the app
  function handleStorageEvent(event: StorageEvent) {
    if (event.key === "index_command_message") {
      try {
        const messageData = event.newValue ? JSON.parse(event.newValue) : null;
        if (messageData) {
          // Log the message and show a notification
          console.log("[index.tsx] Storage event received:", messageData);
          if (Office.context.mailbox && Office.context.mailbox.item) {
            Office.context.mailbox.item.notificationMessages.replaceAsync(
              "IndexCommandMessageNotification",
              {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: `Index received: ${messageData.subject}`,
                icon: "Icon.80x80",
                persistent: false,
              }
            );
          }
          // Write the same message to check the mode
          const item = Office.context.mailbox.item;
          console.log("OnNewMessageComposeHandler: Setting custom property for taskpane", item);
          // Test for read mode
          if (typeof item.subject === "string") {
            // If the subject is a string, we are in read mode.
            console.log("003 Item subject (read mode): " + item.subject);
          } else {
            (item.subject as any).getAsync(function (asyncResult) {
              console.log("003 Item subject (compose mode): " + asyncResult.value);
            });
          }
          // We are in compose mode
        }

        // Clear the item to avoid re-processing
        localStorage.removeItem("index_command_message");
        // after doing this let's reload task pane and re-render
        // This reload doesnot change the Office.context.mailbox.item
        // window.location.reload();
      } catch (error) {
        console.error("[index.tsx] Could not parse message from localStorage: ", error);
      }
    }
  }
  window.addEventListener("storage", handleStorageEvent);

  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
