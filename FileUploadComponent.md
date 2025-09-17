# FileUploadComponentSP — Detailed Explanation

## What this field is
A **SharePoint-aware file picker** built with Fluent UI v9 that behaves like your other DynamicForm fields. It:

- Honors org-wide rules (required/disabled/hidden).
- Never writes to global state on mount.
- Writes **only** when the user acts (pick/remove/clear).
- In **Edit/View**, it can *display* existing SharePoint attachments (read-only).
- For **multi-file** fields, it supports **Add more files** and appends.
- For **single-file** fields (i.e., `multiple=false` or `maxFiles===1`), it replaces.

---

## Inputs (props it accepts)
- `id`: the key used when writing into global form data (e.g., `"attachments"`).
- `displayName`: label shown above the field.
- `multiple`: whether the user can select multiple files at once.
- `accept`: native browser filter (e.g., `.pdf,image/*`).
- `maxFileSizeMB`: per-file size limit.
- `maxFiles`: cap on total number of files when `multiple` is allowed.
- `isRequired`: at least one file must be selected.
- `description`: helper text under the field.
- `className`: layout hook.
- `submitting`: disable the control while the parent is saving.

> **Note:** Networking/context are not passed as props.  
> The component imports your `getFetchAPI` directly and reads `FormCustomizerContext` (SPFx context instance) out of `DynamicFormContext`.

---

## What it reads from the global form context
From `DynamicFormContext` (safely, without assuming full shape):

- **`FormMode`** — convention: `8 = NEW`, `4 = VIEW`.
- **`FormData`** — the item’s data; only `attachments` (boolean) is used.
- **`GlobalFormData(id, value)`** — how the field publishes the user’s selection.
- **`GlobalErrorHandle(id, message | null)`** — how the field reports validation errors.
- **`FormCustomizerContext`** — used to build the REST URL (list title + item ID) for reading existing attachments.

---

## How it renders

### A) NEW mode (`FormMode === 8`)
- **No SharePoint calls.**
- Shows:
  - Primary button to open the file picker.
  - “Clear” button (after selecting files).
  - Hints (accept, max file size, max files).
  - List of selected files (name, type, size) with per-file “remove”.

### B) EDIT/VIEW (`FormMode !== 8`)
- If `FormData.attachments === true`:  
  - Makes **one GET** to `/_api/web/.../AttachmentFiles` and displays existing attachments (icon + link).
  - Existing files are **read-only**, not merged into the selection.
- If `attachments !== true`:  
  - No SharePoint call; shows no “existing attachments” area.
- Then shows the same picker UI as NEW mode.

> Disabled/Hidden flags are respected exactly like other fields.

---

## The “Choose/Add” button logic
We compute:

```ts
isSingleSelection = (!multiple) || (maxFiles === 1);
