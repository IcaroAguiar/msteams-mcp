# [SPIKE] Chat Notes and conversation tabs

> **Status: implemented** as `teams_get_tabs` and `teams_get_notes` (read-only). Live-tested against a work tenant on 2026-11-25 with the Teams client's own Graph token (client `5e3ce6c0`).

**Question:** a group chat's "Notes" tab holds text people type in Teams. `teams_get_thread` shows only a `ThreadActivity/FileAdd` event, `teams_get_shared_files` does not list it, and `teams_get_message` on the event shows no URL. Where does the text live, and can the existing session read it?

## Where chat Notes live

The Notes tab is a **Loop page**, `loop_loopThread.loop`, stored in a **SharePoint Embedded** container (`driveType: other`, description "container created for storing Lightning Loop files", name `IC3_TeamsContainerType_<containerTypeId>`), inside a folder named after the chat (`19_<id>_thread_v2`). It is not a Graph tab and not in the creator's OneDrive.

Teams records it with a chatsvc event in the chat itself:

```xml
<fileadd><eventtime>…</eventtime><initiator>8:orgid:…</initiator>
  <target><id>{guid}</id>
    <uri>https://{tenant}.sharepoint.com/contentstorage/CSP_{guid}/Document%20Library/19_{id}_thread_v2/loop_loopThread.loop?nav=…</uri>
    <type>loop</type><subtype>loopthread</subtype>
    <creator>8:orgid:…</creator><createdAt>{ms}</createdAt>
  </target></fileadd>
```

`teams_get_message` hides this payload because `parseSystemEvent` summarises system events. The `nav` parameter is base64 of `s=…&d={driveId}&f={itemId}&…`.

## What was tried

| Approach | Result |
|---|---|
| Graph `GET /chats/{id}/tabs` | 200, but the Notes page is **not** listed (empty for a chat with Notes). Lists only app tabs. |
| Graph `GET /chats/{id}` | 403: needs `Chat.ReadBasic`/`Chat.Read`, absent from the Teams client's Graph token. |
| chatsvc conversation / thread properties | No Notes or tab property. |
| `GET /storage/fileStorage/containers?$filter=containerTypeId eq …` | 403: the container cannot be discovered without a link. |
| `GET /drives/{containerDriveId}/root/children` | 200, but shows only the folders of chats the user is in; needs the drive ID first. |
| **chatsvc `ThreadActivity/FileAdd` event → `target.uri`** | **Works.** Found by scanning the chat history newest-first. |
| Graph `GET /shares/u!{base64url(uri)}/driveItem` | **200** for the SharePoint Embedded URL (same path `teams_download_file` already uses). |
| Graph `GET /drives/{d}/items/{i}/content?format=html` | **200, `text/html`, `GENERATOR: Microsoft Loop`**: the full page as simple HTML (headings, lists, @mention chips as `mailto:` links). |
| `…/content?format=pdf` | 200 as well; not needed. |
| Raw `.loop` download | Fluid binary snapshot; not usable as text. |

The Graph token used has `TeamsTab.ReadWrite.All`, `Files.ReadWrite.All`, `Sites.ReadWrite.All` and `Notes.ReadWrite.All`; only reads are made.

## Channel tabs

`GET /teams/{groupId}/channels/{channelId}/tabs?$expand=teamsApp` (200) returns the classic tabs. The `groupId` comes from chatsvc `threadProperties.groupId`. Seen in a live channel:

| Tab | `teamsApp.id` | `contentUrl` | Readable as text |
|---|---|---|---|
| Notes (OneNote) | `0d820ecd-def2-4297-adad-78056cde7c78` | `https://www.onenote.com/teams/ChannelNotes?…` (templated) | Not implemented |
| Page (Channel Pages, Loop) | `5ae80e49-7ada-461a-a6bd-c5df2e0cdb06` | SharePoint `/:fl:/r/sites/…/X.loop?…` | Yes, same `/shares` + `format=html` path |
| Files | `3ed5b337-c2c9-4d5d-b7b4-84ff09a8fc1c` | templated file browser | Use `teams_get_shared_files` |

## OneNote

`GET /groups/{groupId}/onenote/pages` answers 200, so the team notebook is reachable with `Notes.ReadWrite.All`. Mapping a channel's OneNote tab to its section (by channel name) was not implemented; `sections?$expand=parentNotebook` returned 500. The tool reports these tabs with their link and the reason.

## Design

- `teams_get_tabs`: Graph app tabs (channel or chat) plus, for chats, a synthetic `notes` tab from the newest `loopthread` FileAdd event in the newest 2,000 messages (`notesSearch` reports the depth). A Graph failure goes to `tabsError` and does not hide Notes.
- `teams_get_notes`: for each readable tab (or one `tabId`, or a `url`), resolve the SharePoint URL with `/shares`, check it is a `.loop`/`.fluid` file, fetch `content?format=html`, flatten with `loopHtmlToText()`.
- Auth failures map away from `AUTH_*` codes, as for downloads, so Graph can never trigger auto-login.

## Open

- A Notes page created before the newest 2,000 messages is not found. The container's drive ID, once known from any FileAdd event, would allow a direct path lookup (`/drives/{d}/root:/19_{id}_thread_v2/loop_loopThread.loop`), if the container is one per tenant; not verified.
- OneNote channel notebooks (see above).
- Meeting notes ("Meeting notes" Loop in meeting chats) were not tested.
