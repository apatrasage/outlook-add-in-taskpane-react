/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
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
