export type IngestStep = "SELECT" | "PREVIEW" | "RUNNING" | "CANCELLED" | "COMPLETE";
export type IngestMode = "FOLDERS" | "EMAILS";
export type Phase = "FETCHING" | "STORING" | "INDEXING" | "DONE";

export type Folder = {
  id: string;
  displayName: string;
  totalItemCount?: number;
};

export type GraphMessage = {
  id: string;
  subject?: string;
  receivedDateTime?: string;
  webLink?: string;
  bodyPreview?: string;
  from?: { emailAddress?: { address?: string } };
};
