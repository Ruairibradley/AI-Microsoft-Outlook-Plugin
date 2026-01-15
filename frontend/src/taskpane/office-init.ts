declare const Office: any;

export async function officeReady(): Promise<void> {
  // When running in a normal browser (not Outlook), Office may be undefined.
  if (typeof Office === "undefined") return;

  await new Promise<void>((resolve) => {
    Office.onReady(() => resolve());
  });
}
