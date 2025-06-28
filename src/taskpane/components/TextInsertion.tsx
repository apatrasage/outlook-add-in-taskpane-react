import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";

/* global HTMLTextAreaElement */

interface TextInsertionProps {
  insertText: (text: string) => void;
}

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [text, setText] = useState<string>("Some text.");

  const handleTextInsertion = async () => {
    await props.insertText(text);
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const handleSendToCommand = () => {
    try {
      // Remove any existing message
      localStorage.removeItem("reverse_command_message");
      localStorage.setItem("reverse_command_message", JSON.stringify({ text }));
      // Optionally, show a notification or feedback
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error("Could not write to localStorage: ", error);
    }
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field
        className={styles.textAreaField}
        size="large"
        label="Enter text to be inserted into the document."
      >
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      <Field className={styles.instructions}>Click the button to insert text.</Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextInsertion}>
        Insert text
      </Button>
      <Button
        appearance="secondary"
        disabled={false}
        size="large"
        onClick={handleSendToCommand}
        style={{ marginTop: 8 }}
      >
        Send to Command (Append to Email)
      </Button>
    </div>
  );
};

export default TextInsertion;
