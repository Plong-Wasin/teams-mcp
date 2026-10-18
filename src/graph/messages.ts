/**
 * The lean shape every message-returning tool projects onto. Raw Graph protocol
 * fields (etag, policyViolation, reactions, etc.) are dropped; body stays as
 * unconverted HTML so it can be read and written back by the edit tools unchanged.
 * Attachments are kept because a Card-only message has an almost empty body.
 */
export interface LeanAttachment {
  id: string | null;
  contentType: string | null;
  name: string | null;
  /** Set for file references. Never downloaded. */
  contentUrl: string | null;
  /** Card payload, parsed from Graph's JSON string; the original string if it does not parse. */
  content?: unknown;
}

export interface LeanMessage {
  id: string;
  /** The Root Message this belongs to, or null if this message IS a Root Message. */
  parentId: string | null;
  from: string | null;
  createdDateTime: string;
  subject: string | null;
  body: string;
  attachments: LeanAttachment[];
  webUrl: string;
}

export function projectAttachment(raw: any): LeanAttachment {
  const lean: LeanAttachment = {
    id: raw.id ?? null,
    contentType: raw.contentType ?? null,
    name: raw.name ?? null,
    contentUrl: raw.contentUrl ?? null,
  };
  if (raw.content != null && raw.content !== "") {
    if (typeof raw.content === "string") {
      try {
        lean.content = JSON.parse(raw.content);
      } catch {
        lean.content = raw.content;
      }
    } else {
      lean.content = raw.content;
    }
  }
  return lean;
}

export function projectMessage(raw: any): LeanMessage {
  return {
    id: raw.id,
    parentId: raw.replyToId ?? null,
    from: raw.from?.user?.displayName ?? null,
    createdDateTime: raw.createdDateTime,
    subject: raw.subject ?? null,
    body: raw.body?.content ?? "",
    attachments: (raw.attachments ?? []).map(projectAttachment),
    webUrl: raw.webUrl,
  };
}

export function sortByCreatedAscending<T extends { createdDateTime: string }>(messages: T[]): T[] {
  return [...messages].sort((a, b) => a.createdDateTime.localeCompare(b.createdDateTime));
}

export const REPLY_PREVIEW_SIZE = 5;
