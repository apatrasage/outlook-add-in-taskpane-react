# Office Add-in Modal/Dialog Prompt: invoiceDetails

## Purpose

Create a direct modal/dialog for displaying invoice details in an Office Add-in using the Office Dialog API, following best practices from https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins.

## Requirements

- **Name:** invoiceDetails
- **Type:** modal/dialog from (https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins)
- **Trigger:** Opened from the `onOpenInvoiceDetailsClick` handler in the codebase
- **API:** Use `Office.context.ui.displayDialogAsync` to open the dialog
- **Entry Point:** Use a dedicated HTML/TSX entry point for the dialog (e.g., `invoiceDetails.html` and `invoiceDetails.tsx`)
- **API Support:** Always check for Dialog API support with `Office.context.requirements.isSetSupported('DialogApi', '1.1')`
- **Events:** Handle dialog events (message, close, error) using the returned dialog object
- **Security:** Dialog content must be served over HTTPS and follow Office Add-in security and accessibility guidelines
- **Accessibility:** Ensure the dialog is accessible (keyboard, screen reader, ARIA, etc.)
- **Build:** The dialog should have its own build entry (like login)
- **Manifest:** Add a manifest entry for the dialog URL
- **Context Passing:** The dialog must be able to make authenticated API calls. Do NOT pass `groupKey`, authentication details, or sensitive context in the URL or query parameters. Instead, use Office Dialog API's event-based communication (e.g., `dialog.messageParent` and `dialog.addEventHandler`) to securely send the context from the parent to the dialog after it is opened.
- **Provider Usage:** In `invoiceDetails.tsx`, receive the context via Office Dialog API message events and provide it to the React component tree using a Context Provider (e.g., `AppContextProvider`). All child components should consume context via React context APIs for API calls and business logic.
- **API Usage Example:** After opening the dialog, use `dialog.messageParent` in the parent to send the full context (e.g., `groupKey`, `token`, user info`). In the dialog, listen for the message event, extract the context, and provide it via context.
- **Security:** Never expose sensitive information in logs or error messages. Sanitize and validate all data received from Office APIs before use.

## Implementation Steps

1. **Create dialog entry files:**
   - `src/apps/invoiceDetails/invoiceDetails.html`
   - `src/apps/invoiceDetails/invoiceDetails.tsx`
2. **Add a build entry for the dialog in `webpack.config.js`**
3. **Add a manifest entry for the dialog URL**
4. **Implement the dialog logic in `invoiceDetails.tsx`**
5. **Open the dialog from `onOpenInvoiceDetailsClick` using the Office Dialog API**
6. **Test for accessibility and Office client compatibility**

## References

- [Office Dialog API documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins)
- [Fluent UI React v9 accessibility](https://react.fluentui.dev/?path=/docs/concepts-accessibility--docs)

---

This prompt ensures the invoiceDetails modal/dialog is implemented in a compliant, accessible, and maintainable way for Office Add-ins.
