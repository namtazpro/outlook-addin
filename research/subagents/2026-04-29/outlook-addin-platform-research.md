<!-- markdownlint-disable-file -->
# Outlook Add-in Platform Research — Read-Mode Mail Add-in

**Date:** 2026-04-29
**Status:** Complete
**Scope:** Research the Office Add-ins (Outlook) platform for building a read-mode mail add-in that surfaces a button while a user is reading an email.

---

## 1. Office.js / Office Add-ins fundamentals

### Manifest types

Outlook add-ins ship with one of two manifest formats:

- **Add-in only manifest (XML)** — the original `<OfficeApp>` XML manifest. Currently the only manifest supported on **Outlook on Mac** and **Outlook on mobile (iOS/Android)**. Schema is documented under the OWEMXML schemas. Version overrides are layered with `<VersionOverrides xsi:type="VersionOverridesV1_0">` and `VersionOverridesV1_1` for Outlook to declare ribbon commands, function files, and event handlers.
- **Unified manifest for Microsoft 365 (JSON)** — same schema used by Teams apps. Single `MicrosoftTeams.schema.json`. Top-level `"extensions"` property maps to `<VersionOverrides>` (ribbons, runtimes, autoRunEvents, alternates, keyboardShortcuts). **NOT supported on Outlook for Mac or Outlook mobile** as of 2026-04 — devs targeting those clients must maintain a parallel XML manifest.

### Command buttons and surfaces (read mode)

For a read-mode button on a message:

- **XML manifest:** declare a `<DesktopFormFactor>` → `<ExtensionPoint xsi:type="MessageReadCommandSurface">` with a `<CustomTab>` or `<OfficeTab id="TabDefault">`, then a `<Group>` containing a `<Control xsi:type="Button">`. The button's `<Action>` is either:
  - `xsi:type="ShowTaskpane"` → opens a task pane (HTML/JS web app).
  - `xsi:type="ExecuteFunction"` → runs a JS function in a hidden runtime; the function must call `event.completed()` and be registered via `Office.actions.associate("functionName", fn)`. JS file is referenced via `<FunctionFile resid="..."/>` pointing to an HTML file (you cannot link a `.js` directly).
- **Mobile XML:** use `<MobileFormFactor>` → `<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">` with a `mobileButton` control.
- **Unified JSON:** declare `"extensions.ribbons"` array entry with `"contexts": ["mailRead"]` (and `"mobile"` formFactor if mobile), `"tabs"` containing groups and controls (`type: "menu" | "button"`), each control's `actionId` matches a runtime-registered function.

### Task pane runtime vs function command runtime

- **Task pane runtime** — hosts your full HTML/JS UI inside a webview/iframe. Persists while the pane is open; full DOM and Office.js are available. UI APIs (`displayDialogAsync`, `messageParent`) work.
- **Function command runtime** — short-lived JS-only runtime triggered by `ExecuteFunction` or events. Imports are not supported in classic Outlook on Windows; bundle code into a single file. UI-mutating APIs are restricted (no dialogs, no `displayMessageForm`, no `getAccessToken` in some Outlook builds — use `OfficeRuntime.auth.getAccessToken`).

### Event-based activation

Declared via `<ExtensionPoint xsi:type="LaunchEvent">` (XML) or `"extensions.autoRunEvents"` (JSON). Handlers must call `event.completed()` within ~300 seconds. Up to 5 event-based add-ins can run concurrently. Production deployment requires admin upload via Microsoft 365 admin center (Integrated apps) or restricted/unrestricted Marketplace listing.

There is **no general "OnMessageRead" event** in production. Two **preview** events exist (classic Outlook on Windows beta only): `OnMessageReadWithCustomAttachment` and `OnMessageReadWithCustomHeader` — they fire when opening a message that matches a manifest-declared attachment type or internet header. For most read-mode scenarios you place a ribbon button (`MessageReadCommandSurface`); event-based activation is primarily a **compose/send** feature (OnNewMessageCompose, OnMessageSend, OnMessageRecipientsChanged, OnMessageAttachmentsChanged, OnMessageFromChanged, etc.).

### Requirement sets

`Mailbox 1.1` through `Mailbox 1.15` are the relevant Outlook requirement sets. Key thresholds for a read-mode add-in:

- 1.1 — base read-mode APIs (attachments, body, itemId, from/to/cc, subject, internetMessageId).
- 1.5 — `getCallbackTokenAsync({isRest:true})`.
- 1.8 — `getAttachmentContentAsync`, `getAttachmentsAsync`, `getAllInternetHeadersAsync`, `getItemIdAsync`, shared folder support.
- 1.9 — `displayReplyFormAsync`/`displayReplyAllFormAsync`.
- 1.10 — `InfobarClicked`, event-based activation introduced.
- 1.13 — shared mailbox support, sensitivity labels.
- 1.14 — `getAsFileAsync` (full MIME of the current message), `closeAsync`.
- 1.15 — `sendAsync`.
- Mobile clients historically capped at 1.5 but additional APIs (e.g., `OnNewMessageCompose`) have been backported.

---

## 2. Cross-client support matrix (2025–2026)

| Feature | Outlook on the web | new Outlook on Windows | classic Outlook on Windows | Outlook on Mac | Outlook on iOS | Outlook on Android |
|---|---|---|---|---|---|---|
| Task pane add-ins | Yes | Yes | Yes (2016+) | Yes | Yes (Read mode only*) | Yes (Read mode only*) |
| Function commands (`ExecuteFunction`) | Yes | Yes | Yes | Yes | Yes (limited) | Yes (limited) |
| `MessageReadCommandSurface` ribbon button | Yes | Yes | Yes | Yes | via `MobileMessageReadCommandSurface` | via `MobileMessageReadCommandSurface` |
| Event-based activation (compose events 1.10+) | Yes | Yes | Yes (Win10 1903+) | Yes (new Mac UI) | Yes (subset: OnNewMessageCompose, OnMessageRecipientsChanged, OnMessageFromChanged) | Same as iOS |
| `OnMessageSend` / Smart Alerts (1.12) | Yes | Yes | Yes | Yes (new Mac UI) | No | No |
| `OnMessageReadWithCustom*` (preview) | No | No | Yes (beta channel) | No | No | No |
| Unified manifest for Microsoft 365 (JSON) | Yes | Yes | Yes | **No** — XML required | **No** — XML required | **No** — XML required |
| IRM-protected items | Yes | Yes | Yes (2009 build+) | Yes (16.77+) | No | No |
| Non-Microsoft accounts (Gmail/Yahoo) | No | No | No | IMAP CloudCache only | No | No |
| COM/VSTO add-ins coexist | n/a | **Not supported** | Yes (avoid same surface) | n/a | n/a | n/a |

`*` Mobile exceptions: online-meeting providers may use `MobileOnlineMeetingCommandSurface` (Appointment Organizer); CRM/note-taking partners may use `MobileLogEventAppointmentAttendee`. Mobile only supports Microsoft 365 / Outlook.com accounts (no on-premises Exchange except some legacy iOS scenarios on classic OWA).

**Practical implication:** if you must reach iOS/Android/Mac users, ship the XML manifest. The unified JSON manifest is the strategic direction but not mobile/Mac compatible yet.

---

## 3. Capturing message metadata, body, and attachments

The current item is `Office.context.mailbox.item`. In read mode it surfaces as `Office.MessageRead`.

### Synchronous read-mode properties (Mailbox 1.1+)

```javascript
const item = Office.context.mailbox.item;
const id = item.itemId;                    // EWS-formatted on desktop/web; REST-formatted on mobile
const subject = item.subject;              // string
const from = item.from;                    // EmailAddressDetails { displayName, emailAddress }
const sender = item.sender;                // EmailAddressDetails (different from from for sent-on-behalf)
const to = item.to;                        // EmailAddressDetails[]
const cc = item.cc;                        // EmailAddressDetails[]
const conversationId = item.conversationId;
const internetMessageId = item.internetMessageId;
const dateTimeCreated = item.dateTimeCreated; // Date
const itemClass = item.itemClass;          // e.g., "IPM.Note"
const attachments = item.attachments;      // AttachmentDetails[] (sync in read mode since 1.1)
```

### Body

```javascript
item.body.getAsync(Office.CoercionType.Html, (result) => {
  if (result.status === Office.AsyncResultStatus.Succeeded) {
    const html = result.value; // sanitized HTML
  }
});
// Also supported: Office.CoercionType.Text
// Mailbox 1.3+: item.body.getAsync with a continuation; for very large bodies prefer Graph or getAsFileAsync (MIME).
```

### Attachments

`AttachmentDetails` exposes `id`, `name`, `contentType`, `size` (bytes), `attachmentType` (`file`, `item`, `cloud`), `isInline`.

```javascript
item.getAttachmentContentAsync(attachmentId, (result) => {
  // result.value: { content: string, format: AttachmentContentFormat }
  // format may be Base64, Eml (item attachment as MIME), ICalendar, or Url (cloud)
});
```

Restrictions:

- `getAttachmentContentAsync` requires Mailbox **1.8**. On older clients fall back to Graph or EWS using the converted REST id.
- Practical attachment size cap returned by `getAttachmentContentAsync` is governed by the host: ~25–35 MB (Exchange Online's standard message size limit, configurable per tenant). Larger files fail with `ErrorAttachmentSizeLimit`.
- Single message size limit is normally 25 MB but tenant admins can raise to 150 MB; mobile clients may further restrict.
- **Cloud attachments** (`attachmentType === "cloud"`) — `getAttachmentContentAsync` returns format `Url` containing a OneDrive/SharePoint sharing link, not the file bytes. To download bytes, follow the link via Graph (`/drives/{driveId}/items/{itemId}/content`) using a delegated token.
- **Inline images** — flagged with `isInline === true`. They appear as `cid:` references in the HTML body. Match by `id` to resolve.

### Internet headers and full MIME

- `getAllInternetHeadersAsync` (1.8) — returns the raw `name: value` header block as a string (read mode, message only).
- `getAsFileAsync` (1.14, read mode message only) — returns the entire message as base64-encoded EML/MIME. Useful for forwarding to compliance/AI services that want the full message.

### Tokens

```javascript
Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (r) => {
  const token = r.value; // bearer for outlook.office.com REST or Graph
});

OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true })
  .then(ssoToken => { /* exchange for Graph token via on-behalf-of */ });
```

---

## 4. When Office.js is insufficient — Graph and EWS

Common reasons to leave Office.js:

1. Need full MIME of an arbitrary message (not just current) — Office.js exposes only the current item.
2. Attachment exceeds the platform's `getAttachmentContentAsync` limit, or you need a **streaming** download.
3. Cloud attachment (OneDrive/SharePoint) — Office.js gives a sharing URL; you must call Graph with a delegated token to download the file.
4. Need to enumerate other folders, search, or write to other items.
5. Need internet headers structured (parsed) — Graph returns them as an array.
6. Backend processing (server-side analysis) — easier with a Graph token issued to the back end.

### Microsoft Graph (recommended path in 2025+)

1. Get an SSO token via `OfficeRuntime.auth.getAccessToken()` (or MSAL Nested App Authentication for newer clients). Configure SSO in the manifest (`webApplicationInfo` / `<WebApplicationInfo>`) and grant the required Graph scopes (`Mail.Read`, `Mail.ReadWrite`, `User.Read`).
2. On a back end, exchange the SSO token using OAuth 2.0 **on-behalf-of** flow for a Graph token (`https://graph.microsoft.com/.default`). Modern guidance favors **Nested App Authentication (NAA)** with MSAL.js so the Graph token can be acquired client-side without OBO.
3. Convert the EWS item id: `Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0)` (skip on mobile — already REST-formatted).
4. Call Graph: `GET https://graph.microsoft.com/v1.0/me/messages/{restId}` (full message); `/messages/{restId}/$value` (raw MIME); `/messages/{restId}/attachments` and `/attachments/{aid}/$value` (binary).

> **Important:** the legacy Outlook REST API (`outlook.office.com/api/v2.0`) and Exchange user-identity tokens are **deprecated**. Extended support for tenant-issued REST tokens ended Oct 14 2025. New work must use Graph + SSO/NAA.

### EWS (`makeEwsRequestAsync`)

- Available only in **classic Outlook on Windows, Mac, and Outlook on the web with Exchange on-prem**. Not available in new Outlook on Windows or on mobile.
- Useful for Exchange on-premises customers without Graph access.
- Send a SOAP envelope; size limit ~1 MB request, ~1 MB response by default.
- Requires `ReadWriteMailbox` permission for most useful operations.
- Considered legacy: prefer Graph wherever possible.

### Permission scopes in the manifest

- XML: `<Permissions>ReadItem | ReadWriteItem | ReadWriteMailbox</Permissions>`.
- Unified JSON: `"extensions.authorization.permissions.resourceSpecific"` for Graph; the add-in-only manifest equivalent is `<WebApplicationInfo>` with `<Scopes>`.

---

## 5. Key references

### Microsoft Learn

- Outlook add-ins overview — https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-add-ins-overview
- Add-in commands (XML) — https://learn.microsoft.com/en-us/office/dev/add-ins/develop/create-addin-commands
- Add-in commands (unified manifest) — https://learn.microsoft.com/en-us/office/dev/add-ins/develop/create-addin-commands-unified-manifest
- Office.context.mailbox.item object model — https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-item-object-model
- Outlook requirement sets — https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets
- Event-based activation — https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation
- Get attachments from Exchange (Graph path) — https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item
- Use Microsoft Graph from an Outlook add-in — https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/microsoft-graph
- (Deprecated) Use Outlook REST APIs — https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/use-rest-api
- Compare XML vs unified manifest — https://learn.microsoft.com/en-us/office/dev/add-ins/develop/json-manifest-overview
- Unified manifest overview / client support — https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview
- Outlook mobile add-ins — https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-mobile-addins
- Add support for add-in commands on mobile — https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/add-mobile-support
- Outlook mobile supported APIs — https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-mobile-apis
- New Outlook on Windows guidance — https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook
- Nested App Authentication (NAA) — https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- Outlook quick start (Yo Office) — https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart-yo

### GitHub samples (OfficeDev)

- Office Add-in samples index — https://github.com/OfficeDev/Office-Add-in-samples
- Outlook Add-in command demo — https://github.com/OfficeDev/outlook-add-in-command-demo
- Outlook SSO with Graph (Mail.Read) — https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO
- Outlook NAA + MSAL.js — https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA
- Set signature (event-based) — https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature
- Smart Alerts (OnMessageSend) — https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories
- Tag external recipients (event-based) — https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external
- Encrypt attachments — https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments
- Verify sensitivity label — https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-verify-sensitivity-label
- Yo Office generator — https://github.com/OfficeDev/generator-office

---

## 6. Top discoveries

1. **No production "OnMessageRead" event.** A read-mode trigger is implemented as a **ribbon button on `MessageReadCommandSurface`**. Auto-launch on read exists only as preview events (`OnMessageReadWithCustomAttachment`/`Header`) on classic Outlook for Windows beta channel.
2. **Choose XML manifest for maximum reach.** Unified JSON manifest is the strategic format but is not yet supported on Outlook for Mac, iOS, or Android — XML is required to ship to those clients.
3. **`MessageReadCommandSurface` button can be `ShowTaskpane` or `ExecuteFunction`.** Function commands run in a JS-only runtime with restricted UI APIs and require `event.completed()` plus `Office.actions.associate(...)`.
4. **Read-mode metadata is largely synchronous.** `subject`, `from`, `to`, `cc`, `itemId`, `internetMessageId`, `attachments`, `conversationId`, `dateTimeCreated` are direct properties on `Office.context.mailbox.item`. Body and attachment content require async calls.
5. **Body retrieval supports HTML or plain text** via `body.getAsync(Office.CoercionType.Html | .Text)`. For full MIME use `getAsFileAsync` (1.14, message read mode).
6. **Attachments expose `attachmentType` of `file`, `item`, or `cloud`.** Cloud attachments do not return bytes via `getAttachmentContentAsync` — they return a OneDrive/SharePoint sharing URL that must be downloaded via Graph with a delegated token.
7. **Attachment / message size practical cap is ~25 MB** for Office.js paths; tenant policy may raise this. Larger or streaming downloads should use Graph `/attachments/{id}/$value` directly.
8. **Outlook REST v2.0 endpoint and legacy Exchange identity tokens are deprecated** (extended support ended October 14 2025). Use Microsoft Graph with SSO via `OfficeRuntime.auth.getAccessToken` or MSAL Nested App Authentication.
9. **EWS via `makeEwsRequestAsync` still works only on classic Outlook for Windows, Mac, and OWA against Exchange on-premises.** It's the fallback for on-prem customers; otherwise prefer Graph.
10. **Item id must be converted before Graph calls** with `Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0)` on every client except mobile (where it is already REST-formatted).
11. **Add-ins do not activate on items opened from `.msg`/`.eml` files, in shared/group/public folders (with limited 1.8/1.13 exceptions), on IRM-protected items in mobile, on delivery reports, or on Simple-MAPI compose windows.** Plan for graceful absence.
12. **Event-based add-ins must be admin-deployed** (or Marketplace-listed with specific options) to auto-launch in production; sideloaded versions do not auto-trigger.

---

## 7. Recommended next research

- [ ] Concrete sample XML manifest for a `MessageReadCommandSurface` button + task pane (verify against current `VersionOverridesV1_1` schema).
- [ ] Concrete unified-manifest equivalent (`extensions.ribbons.contexts: ["mailRead"]`) and the matching `runtimes` entry.
- [ ] NAA/MSAL.js sample wiring for retrieving a Graph token directly from the add-in (avoiding OBO back end).
- [ ] Specific Graph endpoints/queries for downloading file vs cloud attachments and for converting cloud links to drive items.
- [ ] Sideloading and developer testing flow for new Outlook on Windows specifically.
- [ ] Marketplace validation policy for Outlook add-ins claiming mobile support (UI guidelines, Apple Developer ID submission).

## 8. Clarifying questions

1. Which Outlook clients must the read-mode add-in support at v1 — desktop only, or also iOS/Android/Mac (this dictates XML vs unified manifest)?
2. Is on-premises Exchange in scope (would require EWS fallback), or is it Microsoft 365 / Outlook.com only?
3. Does the add-in need to download/process attachments (size threshold? cloud attachments?) or only metadata + body?
4. Where will processing happen — fully in the task pane, or call out to a back-end service (which influences the auth strategy: NAA vs SSO+OBO vs anonymous webhook)?
5. Should the button auto-trigger on opening certain messages (custom-attachment/header preview), or always be a manual click?
