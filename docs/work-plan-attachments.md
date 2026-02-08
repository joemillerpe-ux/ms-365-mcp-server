# Work Plan: To-Do Task Attachment Tools Investigation

**Created:** 2026-02-07
**Tested:** 2026-02-08
**Status:** ❌ ABANDONED - API Limitation Discovered

## Overview

Attempted to add attachment support for To-Do tasks to enable listing, downloading, creating, and copying attachments between tasks.

## Outcome

**Tools were implemented, tested, and REMOVED** because Microsoft Graph API does not expose email attachments dragged into Outlook tasks.

## API Limitation Details

| Endpoint | Result |
|----------|--------|
| `/attachments` (list) | Returns empty array even when UI shows attachments |
| `/attachments/{id}` (get) | N/A - no IDs to retrieve |
| `/linkedResources` (list) | Returns empty array |
| `hasAttachments` property | Shows `true` but API returns no data |
| `expand=["attachments"]` | Accepted but returns no attachment data |

### Testing Performed

1. **Existing tasks with email attachments** - Bronze tasks with visible email attachments in To-Do/Outlook UI returned 0 attachments via API
2. **Brand new task with email** - Created task by dragging email, API still returned empty
3. **File attachments (Excel)** - Dragged Excel file directly into task, API still returned empty
4. **linkedResources** - Also returns empty for all attachment types
5. **expand parameter** - `expand=["attachments"]` on list-todo-tasks returns tasks with no attachment data

### Root Cause

**ALL attachments** on Outlook tasks use an internal Outlook storage mechanism that was supported by the **deprecated Outlook Tasks API** (stopped Aug 2022) but is **NOT exposed** in the current To-Do Graph API. This affects:
- Email attachments (itemAttachment)
- File attachments (Excel, PDF, etc.)
- Any file dragged into task in Outlook

## Workaround: Quick Step "Create task with text of message"

Instead of dragging emails as attachments, use Outlook Quick Step to create tasks:

**Setup:**
1. Create Quick Step in Outlook: "Create a task with text of message"
2. This embeds email content into task body (not as attachment)

**Benefits:**
- Email content accessible via Graph API `body` property
- Claude can read and process the email text
- No attachment API dependency

**Limitation:**
- Embedded images/formatting may not transfer perfectly
- Original email link not preserved (but content is)

## Code Changes (Reverted)

The following were implemented then removed:

### endpoints.json (8 entries removed)
- `list-todo-task-attachments`
- `get-todo-task-attachment`
- `create-todo-task-attachment`
- `delete-todo-task-attachment`
- `list-todo-linked-resources`
- `get-todo-linked-resource`
- `create-todo-linked-resource`
- `delete-todo-linked-resource`

### conversion-tools.ts (2 tools removed)
- `copy-todo-task-attachments`
- `copy-todo-linked-resources`

Comment block added at end of file documenting the limitation.

## API References (for future investigation)

- [List taskFileAttachments](https://learn.microsoft.com/en-us/graph/api/todotask-list-attachments?view=graph-rest-1.0)
- [Get taskFileAttachment](https://learn.microsoft.com/en-us/graph/api/taskfileattachment-get?view=graph-rest-1.0)
- [Create taskFileAttachment](https://learn.microsoft.com/en-us/graph/api/todotask-post-attachments?view=graph-rest-1.0)
- [Attach files overview](https://learn.microsoft.com/en-us/graph/todo-attachments)

## Lessons Learned

1. `hasAttachments: true` does NOT mean attachments are accessible via API
2. Email attachments on tasks use internal Outlook storage not exposed to Graph
3. Always test with real data before building complex copy/migration tools
4. Quick Step "create task with text" is the best workaround for email→task workflows
