import * as React from "react";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (_props: AppProps) => {
  const styles = useStyles();
  const [bgColor, setBgColor] = React.useState<string>("white");

  React.useEffect(() => {
    const item = Office.context.mailbox.item;
    console.log("OnNewMessageComposeHandler: Setting custom property for taskpane", item);
    // Test for read mode
    if (typeof item.subject === "string") {
      // If the subject is a string, we are in read mode.
      console.log("000 OG Item subject (read mode): " + item.subject);
    } else {
      (item.subject as any).getAsync(function (asyncResult) {
        console.log("000 OG Item subject (compose mode): " + asyncResult.value);
      });
    }
    function handleStorageEvent(event: StorageEvent) {
      console.log("Storage event detected:", event);
      if (event.key === "command_message") {
        try {
          const messageData = event.newValue ? JSON.parse(event.newValue) : null;
          if (messageData) {
            setBgColor("#E6F7FF"); // Change to a light blue
            // Optionally, you can display the messageData in the UI or state
            console.log("Storage event received:", messageData);

            // Write the same message to check the mode
            const item = Office.context.mailbox.item;
            console.log("OnNewMessageComposeHandler: Setting custom property for taskpane", item);
            // Test for read mode
            if (typeof item.subject === "string") {
              // If the subject is a string, we are in read mode.
              console.log("002 Item subject (read mode): " + item.subject);
            }
            // We are in compose mode
            item.subject.getAsync(function (asyncResult) {
              console.log("002 Item subject (compose mode): " + asyncResult.value);
            });
          } else {
            console.log("Received an empty message.");
          }
          // Clear the item to avoid re-processing
          localStorage.removeItem("command_message");
        } catch (error) {
          console.error("Could not parse message from localStorage: ", error);
        }
      }
    }

    Office.onReady((info) => {
      if (info.host === Office.HostType.Outlook) {
        console.log("Taskpane is ready and listening for commands.");
        window.addEventListener("storage", handleStorageEvent);
      }
    });

    return () => {
      window.removeEventListener("storage", handleStorageEvent);
    };
  }, []);

  React.useEffect(() => {
    function onItemChanged(_eventArgs: any) {
      // Re-initialize your UI/context here
      const item = Office.context.mailbox.item;
      if (typeof item.subject === "string") {
        // Read mode
        console.log("00A Item subject (read mode):", item.subject);
      } else {
        // Compose mode
        (item.subject as any).getAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("00A Item subject (compose mode):", asyncResult.value);
          }
        });
      }
    }

    if (Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
    }
  }, []);

  return (
    <div className={styles.root} style={{ backgroundColor: bgColor }}>
      <TextInsertion insertText={insertText} />
    </div>
  );
};

export default App;
