/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// Helper function to determine if the current item is a reply/forward
const isReplyOrForward = (): boolean => {
  const item = Office.context.mailbox.item;
  // Check if the item is a message and if it has a conversation ID (indicating it's part of a thread)
  return item.itemType === Office.MailboxEnums.ItemType.Message && !!item.conversationId;
};
/* global Office */
const getOutlookItemContext = ():
  | "Read"
  | "New"
  | "ReplyForward"
  | "AppointmentRead"
  | "AppointmentNew"
  | "Unknown" => {
  const item = Office.context.mailbox.item;

  if (!item) {
    console.error("Office.context.mailbox.item is not available.");
    return "Unknown";
  }

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    // Handle email messages
    if (item.sender) {
      //Correct way to check read mode (if an item has a sender, it means has been sent)
      return "Read";
    }

    // If not sent, it's in compose mode.  Distinguish new from reply/forward.
    if (item.itemClass === "IPM.Note.SMIME" || item.itemClass === "IPM.Note") {
      if (isReplyOrForward()) {
        return "ReplyForward";
      } else {
        return "New";
      }
    } else {
      return "New";
    }
  } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    // Handle appointments
    if (item.organizer) {
      //If has organizer, it means, it is in Read mode
      return "AppointmentRead";
    } else {
      return "AppointmentNew";
    }
  } else {
    return "Unknown";
  }
};

const addHtmlToEmailBody = async (htmlContent: string): Promise<void> => {
  const context = getOutlookItemContext();

  // Ensure the mode is supported (e.g., "New" or "ReplyForward")
  if (context !== "New" && context !== "ReplyForward") {
    throw new Error(
      `Unsupported mode: ${context}. HTML content can only be added in 'New' or 'ReplyForward' modes.`
    );
  }

  // Ensure the Office context is available
  if (Office.context.mailbox.item) {
    // Generate a unique ID for the focus marker
    const uniqueId = `focus-marker-${Date.now()}`;
    const combinedHtmlContent = `
        ${htmlContent}
        <span id="${uniqueId}"></span>
      `;

    // Insert the combined HTML content
    Office.context.mailbox.item.body.setSelectedDataAsync(
      combinedHtmlContent,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          // Restore focus to the unique marker
          setTimeout(() => {
            Office.context.mailbox.item.body.setSelectedDataAsync(
              `<span id="${uniqueId}"></span>`, // Use the original unique marker to set focus
              { coercionType: Office.CoercionType.Html },
              (focusResult) => {
                if (focusResult.status === Office.AsyncResultStatus.Succeeded) {
                } else {
                  console.error(
                    "Failed to restore focus to the email body:",
                    focusResult.error.message
                  );
                }
              }
            );
          }, 100); // Delay to ensure the content is set before restoring focus
        } else {
          console.error("Failed to add HTML content to the email body:", result.error.message);
        }
      }
    );
  } else {
    throw new Error("Office context is not available.");
  }
};

Office.onReady(async () => {
  // If needed, Office.js is ready to be called.
  // Reverse journey: check for message from taskpane
  async function handleReverseCommandMessage(msg: string | null) {
    if (msg) {
      try {
        const { text } = JSON.parse(msg);
        await addHtmlToEmailBody(text);
        localStorage.removeItem("reverse_command_message");
      } catch (err) {
        // eslint-disable-next-line no-console
        console.error("Error handling reverse_command_message:", err);
      }
    }
  }

  // Initial check on load
  await handleReverseCommandMessage(localStorage.getItem("reverse_command_message"));

  // Listen for storage events
  window.addEventListener("storage", (event) => {
    if (event.key === "reverse_command_message" && event.newValue) {
      handleReverseCommandMessage(event.newValue);
    }
  });
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

/**
 * Handler for OnNewMessageCompose event. Shows a notification and completes the event.
 * @param event Office.AddinCommands.Event
 */
function OnNewMessageComposeHandler(event: Office.AddinCommands.Event) {
  // Set a custom property to signal the event to the taskpane
  const item = Office.context.mailbox.item;
  console.log("OnNewMessageComposeHandler: Setting custom property for taskpane", item);
  // Test for read mode
  if (typeof item.subject === "string") {
    // If the subject is a string, we are in read mode.
    console.log("001 Item subject (read mode): " + item.subject);
  }

  // We are in compose mode
  item.subject.getAsync(function (asyncResult) {
    console.log("001 Item subject (compose mode): " + asyncResult.value);
    const messagePayload = {
      subject: asyncResult.value,
      from: item.from ? item.from.emailAddress : undefined,
      timestamp: new Date().toISOString(),
    };
    console.log("messagePayload:", messagePayload);
    try {
      localStorage.setItem("command_message", JSON.stringify(messagePayload));
      localStorage.setItem("index_command_message", JSON.stringify(messagePayload));
    } catch (error) {
      console.error("Could not write to localStorage: ", error);
    }
  });

  event.completed();

  // if (Office.context.mailbox && Office.context.mailbox.item) {
  //   Office.context.mailbox.item.loadCustomPropertiesAsync((asyncResult) => {
  //     if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //       const customProps = asyncResult.value;
  //       customProps.set("OnNewMessageCompose", "1");
  //       console.log("OnNewMessageComposeHandler: Custom property set");
  //       customProps.saveAsync(() => {
  //         // Optionally, show a notification
  //         const message: Office.NotificationMessageDetails = {
  //           type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //           message: "OnNewMessageCompose event triggered!",
  //           icon: "Icon.80x80",
  //           persistent: true,
  //         };
  //         Office.context.mailbox.item.notificationMessages.replaceAsync(
  //           "OnNewMessageComposeNotification",
  //           message,
  //           () => {
  //             console.log("OnNewMessageComposeHandler: Notification shown, event completed");
  //             event.completed();
  //           }
  //         );
  //       });
  //     } else {
  //       console.log("OnNewMessageComposeHandler: Failed to load custom properties");
  //       event.completed();
  //     }
  //   });
  // } else {
  //   console.log("OnNewMessageComposeHandler: No mailbox item available");
  //   event.completed();
  // }
}

// Register the function with Office.
Office.actions.associate("action", action);

// Register the function with Office actions
Office.actions.associate("OnNewMessageComposeHandler", OnNewMessageComposeHandler);
