import * as React from "react";
import { Button, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    background: "#f5f5f5",
    padding: "2rem",
  },
  title: {
    fontSize: "1.5rem",
    fontWeight: 600,
    marginBottom: "1rem",
  },
  content: {
    marginBottom: "2rem",
    color: "#333",
  },
});

const DialogApp: React.FC = () => {
  const styles = useStyles();

  const handleClose = () => {
    if (window.Office && window.Office.context && window.Office.context.ui) {
      window.Office.context.ui.messageParent("Dialog closed by user");
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.title}>Sample Dialog</div>
      <div className={styles.content}>
        This is a dummy dialog content rendered with React and TypeScript.
        <br />
        You can close this dialog using the button below.
      </div>
      <Button appearance="primary" onClick={handleClose} aria-label="Close dialog">
        Close Dialog
      </Button>
    </div>
  );
};

export default DialogApp;
