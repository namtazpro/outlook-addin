/* global Office */

export interface EmailAddress {
  name?: string;
  address: string;
}

export interface EmailAttachment {
  id: string;
  name: string;
  size: number;
  contentType?: string;
  isInline: boolean;
  contentBase64: string | null;
  error?: string;
}

export interface ProjectRef {
  id: string;
  name: string;
}

export interface EmailPayload {
  internetMessageId: string;
  subject: string;
  sentDate: string | null;
  from: EmailAddress | null;
  to: EmailAddress[];
  cc: EmailAddress[];
  bodyText: string;
  attachments: EmailAttachment[];
  capturedAt: string;
  capturedBy: EmailAddress | null;
  project: ProjectRef | null;
}

function toEmailAddress(d: Office.EmailAddressDetails | undefined | null): EmailAddress | null {
  if (!d) return null;
  return { name: d.displayName, address: d.emailAddress };
}

function toEmailAddresses(arr: Office.EmailAddressDetails[] | undefined): EmailAddress[] {
  if (!arr) return [];
  return arr.map((d) => ({ name: d.displayName, address: d.emailAddress }));
}

function getBodyText(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value ?? "");
      } else {
        reject(new Error(result.error?.message ?? "Failed to read body"));
      }
    });
  });
}

function getAttachmentBase64(item: Office.MessageRead, attachmentId: string): Promise<string> {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(result.error?.message ?? "Failed to read attachment"));
        return;
      }
      const value = result.value;
      if (value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
        resolve(value.content);
      } else {
        // Eml/iCal/Url formats — wrap as-is in base64 of UTF-8.
        try {
          // btoa is fine here; content is text in those formats.
          resolve(btoa(unescape(encodeURIComponent(value.content))));
        } catch (e) {
          reject(e instanceof Error ? e : new Error(String(e)));
        }
      }
    });
  });
}

export async function buildEmailPayload(project: ProjectRef | null): Promise<EmailPayload> {
  const item = Office.context.mailbox.item as Office.MessageRead | undefined;
  if (!item) {
    throw new Error("No mail item is currently selected.");
  }

  const bodyText = await getBodyText(item);

  const attachmentDetails = item.attachments ?? [];
  const attachments: EmailAttachment[] = await Promise.all(
    attachmentDetails.map(async (a): Promise<EmailAttachment> => {
      try {
        const contentBase64 = await getAttachmentBase64(item, a.id);
        return {
          id: a.id,
          name: a.name,
          size: a.size,
          contentType: a.contentType,
          isInline: a.isInline,
          contentBase64,
        };
      } catch (e) {
        return {
          id: a.id,
          name: a.name,
          size: a.size,
          contentType: a.contentType,
          isInline: a.isInline,
          contentBase64: null,
          error: e instanceof Error ? e.message : String(e),
        };
      }
    })
  );

  const profile = Office.context.mailbox.userProfile;
  const capturedBy: EmailAddress | null = profile
    ? { name: profile.displayName, address: profile.emailAddress }
    : null;

  return {
    internetMessageId: item.internetMessageId ?? "",
    subject: item.subject ?? "",
    sentDate: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
    from: toEmailAddress(item.from),
    to: toEmailAddresses(item.to),
    cc: toEmailAddresses(item.cc),
    bodyText,
    attachments,
    capturedAt: new Date().toISOString(),
    capturedBy,
    project,
  };
}
