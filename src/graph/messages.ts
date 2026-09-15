/**
 * The lean shape every message-returning tool projects onto. Raw Graph protocol
 * fields (etag, policyViolation, reactions, etc.) are dropped; body stays as
 * unconverted HTML so it can be read and written back by the edit tools unchanged.
 */
export interface LeanMessage {
  id: string;
  /** The Root Message this belongs to, or null if this message IS a Root Message. */
  parentId: string | null;
  from: string | null;
  createdDateTime: string;
  subject: string | null;
  body: string;
  webUrl: string;
}

export function projectMessage(raw: any): LeanMessage {
  return {
    id: raw.id,
    parentId: raw.replyToId ?? null,
    from: raw.from?.user?.displayName ?? null,
    createdDateTime: raw.createdDateTime,
    subject: raw.subject ?? null,
    body: raw.body?.content ?? "",
    webUrl: raw.webUrl,
  };
}

export function sortByCreatedAscending<T extends { createdDateTime: string }>(messages: T[]): T[] {
  return [...messages].sort((a, b) => a.createdDateTime.localeCompare(b.createdDateTime));
}

export const REPLY_PREVIEW_SIZE = 5;
