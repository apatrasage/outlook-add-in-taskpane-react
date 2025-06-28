/* global Office console */

export async function insertText(text: string): Promise<void> {
  // Write text to the cursor point in the compose surface.
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item?.body.setSelectedDataAsync(
        text,
        { coercionType: Office.CoercionType.Text },
        (asyncResult: Office.AsyncResult<void>) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(asyncResult.error.message));
          } else {
            resolve();
          }
        }
      );
    } catch (error) {
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });
}
