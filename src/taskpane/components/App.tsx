import * as React from "react";
import TextInsertion from "./TextInsertion";
import { makeStyles, Button, Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  dialogButton: {
    margin: "1rem",
    padding: "0.5rem 1rem",
  },
  rootDynamicBgBlue: {
    minHeight: "100vh",
    backgroundColor: "#E6F7FF",
  },
  rootDynamicBgWhite: {
    minHeight: "100vh",
    backgroundColor: "white",
  },
});

const App: React.FC<AppProps> = (_props: AppProps) => {
  const styles = useStyles();
  const [bgColor, setBgColor] = React.useState<string>("white");
  const [dialogOpen, setDialogOpen] = React.useState<boolean>(false);
  const [errorOpen, setErrorOpen] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string>("");

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

  /**
   * Opens an Office dialog using Office.js Dialog API
   * If error 12007 occurs, show an error dialog with a Reload option; no retries.
   */
  const openDialogWithRetry = React.useCallback((): void => {
    const dialogUrl = `${window.location.origin}/dialog.html`;

    console.log(`Opening dialog at URL: ${dialogUrl}`);

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 40, width: 30, displayInIframe: true },
      (result: Office.AsyncResult<Office.Dialog>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          setDialogOpen(true);

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
            // Handle message from dialog
            if ("message" in arg) {
              console.log("Dialog message received:", (arg as any).message);
            } else {
              console.log("Dialog event received:", arg);
            }
            dialog.close();
            setDialogOpen(false);
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
            // Handle dialog closed or error
            console.log("Dialog event:", event);
            setDialogOpen(false);
          });
        } else {
          // Show error prompt for 12007 (dialog already open); do not retry
          if ((result.error as Office.Error).code === 12007) {
            console.warn("Dialog open failed with error 12007. Showing reload prompt.");
            setDialogOpen(false);
            setErrorMessage(
              "We couldn't open the dialog because another dialog is already open. Please reload the task pane and try again."
            );
            setErrorOpen(true);
          } else {
            // Handle other errors
            console.error("Failed to open dialog:", result.error);
            setDialogOpen(false);
          }
        }
      }
    );
  }, []);

  /**
   * Public method to open dialog - calls the (no-retry) implementation
   */
  const openDialog = React.useCallback((): void => {
    if (dialogOpen) {
      console.warn("Dialog is already open, ignoring request");
      return;
    }

    openDialogWithRetry();
  }, [dialogOpen, openDialogWithRetry]);

  return (
    <div className={bgColor === "#E6F7FF" ? styles.rootDynamicBgBlue : styles.rootDynamicBgWhite}>
      <TextInsertion insertText={insertText} />
      {/* Button to open dialog with improved accessibility */}
      <Button
        appearance="primary"
        onClick={openDialog}
        disabled={dialogOpen}
        aria-label="Open dialog window"
        className={styles.dialogButton}
      >
        {dialogOpen ? "Dialog Open..." : "Open Dialog"}
      </Button>

      {/* Error dialog prompting to reload */}
      <Dialog open={errorOpen} onOpenChange={(_, data) => setErrorOpen(!!data.open)}>
        <DialogSurface aria-describedby="dialog-error-desc">
          <DialogBody>
            <DialogTitle>Problem opening dialog</DialogTitle>
            <DialogContent id="dialog-error-desc">
              {errorMessage || "An unexpected error occurred while opening the dialog."}
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setErrorOpen(false)}>
                Close
              </Button>
              <Button appearance="primary" onClick={() => window.location.reload()}>
                Reload
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default App;
