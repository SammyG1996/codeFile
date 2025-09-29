/**
 * FileUploadComponent.tsx
 * ---------------------------------------------------------------------
 * Purpose
 *  - A reusable file upload field that works inside your DynamicForm/SPFx
 *    Form Customizer solution.
 *  - Supports single or multiple file selection, basic client-side
 *    validation (required, max files, max size), and shows existing
 *    SharePoint attachments when editing/viewing a list item.
 *
 * How to use
 * ---------------------------------------------------------------------
 * <FileUploadComponent
 *   id="Attachments"               // the key used when writing to GlobalFormData
 *   displayName="Attachments"      // label shown in the UI
 *   multiple                       // allow selecting more than one file
 *   accept=".pdf,.doc,.docx,image/*"
 *   maxFiles={5}                   // maximum number of files user can pick
 *   maxFileSizeMB={15}             // per-file size limit in MB
 *   isRequired={false}             // validate that at least one file is selected
 *   description="Add any supporting files."
 *   submitting={isSubmitting}      // disable while the form is saving
 *   context={props.context}        // SPFx FormCustomizerContext for REST URLs
 * />
 *
 * Notes
 * ---------------------------------------------------------------------
 * - This component only SELECTS files and validates them. Actual upload to
 *   SharePoint is expected to happen in your form's submit handler where
 *   you read GlobalFormData(id) and post files as you need.
 * - When in Edit/View mode, if FormData.Attachments indicates the item has
 *   attachments, we call the SharePoint REST API to list them and display
 *   links (read-only).
 */

import * as React from 'react';
import { Field, Button, Text, Link } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';

// Form-level context (your existing provider) used to commit field values/errors
import { DynamicFormContext } from './DynamicFormContext';

// SPFx Form Customizer context gives us list title/GUID, item ID, and web URL
import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

// Project helper for making SharePoint REST requests
import { getFetchAPI } from '../Utilis/getFetchApi';

/* ============================= Types ============================= */

// Props accepted by the component
export interface FileUploadProps {
  id: string;                      // field key used when writing to GlobalFormData
  displayName: string;             // visible label
  multiple?: boolean;              // allow selecting multiple files
  accept?: string;                 // accept attribute (e.g. ".pdf,image/*")
  maxFileSizeMB?: number;          // per-file size limit (MB)
  maxFiles?: number;               // overall limit for file count (only for multi-select)
  isRequired?: boolean;            // require at least one file
  description?: string;            // helper text below the field
  className?: string;              // extra CSS class for action row
  submitting?: boolean;            // disable interactions while submitting
  context?: FormCustomizerContext; // SPFx context (build absolute REST URLs)
}

// Subset of your DynamicFormContext shape we rely on
type FormCtxShape = {
  FormData?: Record<string, unknown>;               // current item data (Edit/View) or defaults (New)
  FormMode?: number;                                // 8=new, 6=edit, 4=view per your solution
  GlobalFormData: (id: string, value: unknown) => void;          // commit value to form
  GlobalErrorHandle: (id: string, error: string | undefined) => void; // commit error to form

  // optional flags/lists handled by your provider
  isDisabled?: boolean;
  disabled?: boolean;
  formDisabled?: boolean;
  Disabled?: boolean;
  AllDisabledFields?: unknown;
  AllHiddenFields?: unknown;
};

// Minimal shape for an attachment returned by REST
type SPAttachment = { FileName: string; ServerRelativeUrl: string };

/* =========================== Utilities =========================== */

// User messages for validation
const REQUIRED_MSG = 'Please select a file.';
const TOO_LARGE_MSG = (name: string, limitMB: number): string =>
  `“${name}” exceeds the maximum size of ${limitMB} MB.`;
const TOO_MANY_MSG = (limit: number): string =>
  `You can attach up to ${limit} file${limit === 1 ? '' : 's'}.`;

// Small helpers
const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

// Read a boolean-ish flag from a bag of possible keys on an unknown object
const getCtxFlag = (o: Record<string, unknown>, keys: string[]): boolean =>
  keys.some(k => Object.prototype.hasOwnProperty.call(o, k) && Boolean(o[k]));

// Check if a display name is present in a list-like "disabled/hidden fields" structure
const isListed = (bag: unknown, name: string): boolean => {
  const needle = name.trim().toLowerCase();
  if (bag === null || bag === undefined) return false;

  if (Array.isArray(bag)) return bag.some(v => String(v).trim().toLowerCase() === needle);

  if (typeof (bag as { has?: unknown }).has === 'function') {
    for (const v of bag as Set<unknown>) if (String(v).trim().toLowerCase() === needle) return true;
    return false;
  }

  if (typeof bag === 'string') return bag.split(',').map(s => s.trim().toLowerCase()).includes(needle);

  if (typeof bag === 'object') {
    for (const [k, v] of Object.entries(bag as Record<string, unknown>)) {
      if (k.trim().toLowerCase() === needle && Boolean(v)) return true;
    }
  }
  return false;
};

// Format file sizes nicely for the list
const formatBytes = (bytes: number): string => {
  if (!Number.isFinite(bytes)) return '';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0;
  let n = bytes;
  while (n >= 1024 && i < units.length - 1) {
    n /= 1024;
    i++;
  }
  return `${Number.isInteger(n) ? n.toFixed(0) : n.toFixed(2)} ${units[i]}`;
};

// Look at FormData and interpret different possible keys that mean "has attachments"
const readAttachmentsHint = (fd: Record<string, unknown> | undefined): boolean | undefined => {
  if (!fd) return undefined;
  const tryKeys = ['Attachments', 'attachments', 'AttachmentCount', 'attachmentCount'] as const;
  for (const k of tryKeys) {
    if (Object.prototype.hasOwnProperty.call(fd, k)) {
      const v = (fd as Record<string, unknown>)[k];
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v > 0;
    }
  }
  return undefined;
};

/* ============================ Component ============================ */

export default function FileUploadComponent(props: FileUploadProps): JSX.Element {
  const {
    id,
    displayName,
    multiple = false,
    accept,
    maxFileSizeMB,
    maxFiles,
    isRequired,
    description = '',
    className,
    submitting,
    context, // SPFx FormCustomizerContext
  } = props;

  // Pull the form-level services + data (commit hooks, mode, optional flags)
  const raw = React.useContext(DynamicFormContext) as unknown as FormCtxShape;

  // Basic mode and existing data
  const FormData = raw.FormData;
  const FormMode = raw.FormMode ?? 0;
  const isDisplayForm = FormMode === 4; // VIEW mode
  const isNewMode = FormMode === 8;     // NEW mode

  // Merge multiple sources that can disable/hide this field (your provider's design)
  const disabledFromCtx = getCtxFlag(raw as unknown as Record<string, unknown>, [
    'isDisabled', 'disabled', 'formDisabled', 'Disabled',
  ]);
  const AllDisabledFields = raw.AllDisabledFields;
  const AllHiddenFields = raw.AllHiddenFields;

  // ========================= Local UI State =========================
  // Whether field is required (prop can change)
  const [required, setRequired] = React.useState<boolean>(Boolean(isRequired));

  // Disable or hide based on mode, provider flags/lists, and submitting
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || Boolean(submitting) || isListed(AllDisabledFields, displayName)
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(isListed(AllHiddenFields, displayName));

  // Files the user picked in this session (not yet uploaded)
  const [files, setFiles] = React.useState<File[]>([]);
  const [error, setError] = React.useState<string>(''); // local validation error message

  // Read-only list of existing SP attachments for this item (Edit/View)
  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | undefined>(undefined);
  const [loadingSP, setLoadingSP] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');

  // Hidden input reference so we can trigger the OS picker with a button
  const inputRef = React.useRef<HTMLInputElement>(null);

  // If multiple=false OR maxFiles===1, we behave as a single-file input
  const isSingleSelection = !multiple || maxFiles === 1;

  /* ------------------------- Effects: props -> state ------------------------- */

  // If the `isRequired` prop changes, update state
  React.useEffect((): void => {
    setRequired(Boolean(isRequired));
  }, [isRequired]);

  // Recompute disabled/hidden when environment changes
  React.useEffect((): void => {
    const fromMode = isDisplayForm;
    const fromCtx = disabledFromCtx;
    const fromSubmitting = Boolean(submitting);
    const fromDisabledList = isListed(AllDisabledFields, displayName);
    const fromHiddenList = isListed(AllHiddenFields, displayName);

    setIsDisabled(fromMode || fromCtx || fromDisabledList || fromSubmitting);
    setIsHidden(fromHiddenList);
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  /* ----------------- Effects: fetch existing attachments ----------------- */
  /**
   * In Edit/View:
   *   If FormData indicates there are attachments on the item,
   *   we call the SharePoint REST API to retrieve AttachmentFiles and display them.
   *
   * We build absolute URLs using:
   *   - context.pageContext.web.absoluteUrl
   *   - context.list.id (GUID) OR context.list.title
   *   - context.item.ID
   */
  React.useEffect((): void | (() => void) => {
    // New mode has no existing attachments
    if (isNewMode) return;

    // Only fetch if the item is expected to have attachments
    const attachmentsHint = readAttachmentsHint(FormData);
    if (attachmentsHint === false) {
      setSpAttachments([]); // explicitly show "No existing attachments"
      setLoadingSP(false);
      setLoadError('');
      return;
    }

    // Safe optional chaining to read parts of the SPFx context
    const listTitle: string | undefined = (context as { list?: { title?: string } } | undefined)?.list?.title;
    const listGuid: string | undefined = (context as { list?: { id?: string } } | undefined)?.list?.id;
    const itemId: number | undefined = (context as { item?: { ID?: number } } | undefined)?.item?.ID;
    const baseUrl: string | undefined =
      (context as { pageContext?: { web?: { absoluteUrl?: string } } } | undefined)?.pageContext?.web?.absoluteUrl ??
      (typeof window !== 'undefined' ? window.location.origin : undefined);

    // Need a base URL, the item ID, and either a GUID or a Title to address the list
    if (!baseUrl || !itemId || (!listGuid && !listTitle)) {
      return;
    }

    // Encode title/ID for the URL paths
    const encTitle = listTitle ? encodeURIComponent(listTitle) : '';
    const idStr = encodeURIComponent(String(itemId));

    // We try a few URL shapes to be robust to helper implementations:
    //  - with "/_api" and without (some helpers prepend it)
    //  - addressing by GUID (preferred) or by Title
    //  - using items(<Id>) and the $filter=Id eq <Id> form
    const urls: string[] = [];
    if (listGuid) {
      urls.push(
        `${baseUrl}/_api/web/lists(guid'${listGuid}')/items(${idStr})?$select=AttachmentFiles&$expand=AttachmentFiles`,
        `${baseUrl}/web/lists(guid'${listGuid}')/items(${idStr})?$select=AttachmentFiles&$expand=AttachmentFiles`,
        `${baseUrl}/_api/web/lists(guid'${listGuid}')/items?$filter=Id eq ${idStr}&$select=AttachmentFiles&$expand=AttachmentFiles`,
        `${baseUrl}/web/lists(guid'${listGuid}')/items?$filter=Id eq ${idStr}&$select=AttachmentFiles&$expand=AttachmentFiles`
      );
    }
    if (listTitle) {
      urls.push(
        `${baseUrl}/_api/web/lists/getbytitle('${encTitle}')/items(${idStr})?$select=AttachmentFiles&$expand=AttachmentFiles`,
        `${baseUrl}/web/lists/getbytitle('${encTitle}')/items(${idStr})?$select=AttachmentFiles&$expand=AttachmentFiles`,
        `${baseUrl}/_api/web/lists/getbytitle('${encTitle}')/items?$filter=Id eq ${idStr}&$select=AttachmentFiles&$expand=AttachmentFiles`,
        `${baseUrl}/web/lists/getbytitle('${encTitle}')/items?$filter=Id eq ${idStr}&$select=AttachmentFiles&$expand=AttachmentFiles`
      );
    }

    let cancelled = false;

    // Async IIFE to run the fetch; we `.catch()` to satisfy ESLint (no-floating-promises)
    (async (): Promise<void> => {
      setLoadingSP(true);
      setLoadError('');

      let success = false;
      let lastErr: unknown = null;

      // Try each candidate URL until one works
      for (const spUrl of urls) {
        if (cancelled) return;

        try {
          // Your project helper; expects an absolute URL string
          const respUnknown: unknown = await getFetchAPI({
            spUrl,
            method: 'GET',
            headers: { Accept: 'application/json;odata=nometadata' },
          });

          // Response can be a single item or a collection with "value"
          // We normalize both shapes to a plain "attachments" array.
          let attsRaw: unknown;
          if (respUnknown && typeof respUnknown === 'object') {
            const r = respUnknown as Record<string, unknown>;
            if (Array.isArray(r.value)) {
              const first = r.value[0] as Record<string, unknown> | undefined;
              attsRaw = first?.AttachmentFiles;
            } else if (Object.prototype.hasOwnProperty.call(r, 'AttachmentFiles')) {
              attsRaw = (r as { AttachmentFiles?: unknown }).AttachmentFiles;
            }
          }

          // Map unknown objects into our SPAttachment type (skip anything malformed)
          const atts: SPAttachment[] = Array.isArray(attsRaw)
            ? (attsRaw as unknown[]).map((x): SPAttachment | undefined => {
                if (x && typeof x === 'object') {
                  const o = x as Record<string, unknown>;
                  const FileName = typeof o.FileName === 'string' ? o.FileName : undefined;
                  const ServerRelativeUrl = typeof o.ServerRelativeUrl === 'string' ? o.ServerRelativeUrl : undefined;
                  if (FileName && ServerRelativeUrl) return { FileName, ServerRelativeUrl };
                }
                return undefined;
              }).filter((x): x is SPAttachment => !!x)
            : [];

          setSpAttachments(atts);
          setLoadingSP(false);
          success = true;
          break; // stop after first success
        } catch (e: unknown) {
          lastErr = e; // remember error but keep trying next URL
        }
      }

      // If none of the candidates worked, surface the error
      if (!success && !cancelled) {
        const msg = lastErr instanceof Error ? lastErr.message : 'Failed to load attachments.';
        setSpAttachments(undefined);
        setLoadError(msg);
        setLoadingSP(false);
      }
    })().catch(() => {
      // Swallowing to satisfy lints; errors handled in the code above.
    });

    // Cleanup if the component unmounts before the fetch completes
    return (): void => {
      cancelled = true;
    };
  }, [isNewMode, FormData, context]);

  /* ------------------------- Validation & commit ------------------------- */

  // Validate picked files against required/maxFiles/maxFileSizeMB
  const validateSelection = React.useCallback(
    (list: File[]): string => {
      if (required && list.length === 0) return REQUIRED_MSG;

      if (!isSingleSelection && isDefined(maxFiles) && list.length > maxFiles) {
        return TOO_MANY_MSG(maxFiles);
      }

      if (isDefined(maxFileSizeMB)) {
        const perFileLimitBytes = maxFileSizeMB * 1024 * 1024;
        for (const f of list) {
          if (f.size > perFileLimitBytes) return TOO_LARGE_MSG(f.name, maxFileSizeMB);
        }
      }
      return '';
    },
    [required, isSingleSelection, maxFiles, maxFileSizeMB]
  );

  // Push the selected files into your form's GlobalFormData store.
  // For single mode we send the single File object; for multi we send the array.
  const commitValue = React.useCallback(
    (list: File[]): void => {
      const payload: unknown = list.length === 0 ? undefined : isSingleSelection ? list[0] : list;
      raw.GlobalFormData(id, payload);
    },
    [raw, id, isSingleSelection]
  );

  /* ----------------------------- Event handlers ----------------------------- */

  // Open the OS file picker (the input is hidden)
  const openPicker = (): void => {
    if (!isDisabled) inputRef.current?.click();
  };

  // When the user selects files in the dialog
  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e): void => {
    const picked = Array.from(e.currentTarget.files ?? []);

    let next: File[] = [];
    let msg = '';

    if (isSingleSelection) {
      // Single mode: take the first file only
      next = picked.slice(0, 1);
    } else {
      // Multi mode: append to any existing selection, respecting maxFiles
      const already = files.length;
      const capacity = isDefined(maxFiles) ? Math.max(0, maxFiles - already) : picked.length;

      if (already === 0) {
        const toTake = isDefined(maxFiles) ? Math.min(picked.length, maxFiles) : picked.length;
        next = picked.slice(0, toTake);
        if (isDefined(maxFiles) && picked.length > maxFiles) msg = TOO_MANY_MSG(maxFiles);
      } else {
        const toAdd = picked.slice(0, capacity);
        next = files.concat(toAdd);
        if (isDefined(maxFiles) && picked.length > capacity) msg = TOO_MANY_MSG(maxFiles);
      }
    }

    // If count looks good, still check per-file size and "required"
    if (!msg) msg = validateSelection(next);

    // Update local state + propagate error/value to the form
    setFiles(next);
    setError(msg);
    raw.GlobalErrorHandle(id, msg === '' ? undefined : msg);
    commitValue(next);

    // Reset the input so the same file can be selected again if needed
    if (inputRef.current) inputRef.current.value = '';
  };

  // Remove a single file by index (for multi-select scenario)
  const removeAt = React.useCallback(
    (idx: number): void => {
      const next = files.filter((_, i) => i !== idx);
      const msg = validateSelection(next);

      setFiles(next);
      setError(msg);
      raw.GlobalErrorHandle(id, msg === '' ? undefined : msg);
      commitValue(next);
    },
    [files, validateSelection, raw, id, commitValue]
  );

  // Wrap removeAt for Button onClick
  const handleRemove = React.useCallback(
    (idx: number): React.MouseEventHandler<HTMLButtonElement> =>
      (): void => removeAt(idx),
    [removeAt]
  );

  // Clear local selection completely
  const clearAll = (): void => {
    const msg = required ? REQUIRED_MSG : '';
    setFiles([]);
    setError(msg);
    raw.GlobalErrorHandle(id, msg || undefined);
    commitValue([]);
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ------------------------------- Render ------------------------------- */

  // If the field is hidden by rule, render nothing (but keep DOM consistent)
  if (isHidden) return <div hidden className="fieldClass" />;

  return (
    <div className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {/* ===== Existing SP attachments (Edit/View only) ===== */}
        {!isNewMode && (
          <div style={{ marginBottom: 8 }}>
            {loadingSP && <Text size={200}>Loading attachments…</Text>}

            {!loadingSP && loadError && (
              <Text size={200} aria-live="polite">
                Error: {loadError}
              </Text>
            )}

            {!loadingSP && !loadError && Array.isArray(spAttachments) && spAttachments.length > 0 && (
              <div style={{ display: 'grid', gap: 6 }}>
                {spAttachments.map((a, i) => (
                  <div
                    key={`${a.ServerRelativeUrl}-${i}`}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: 8,
                      padding: '6px 10px',
                      borderRadius: 8,
                      border: '1px solid var(--colorNeutralStroke1)',
                    }}
                  >
                    <DocumentRegular />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div
                        style={{
                          fontWeight: 500,
                          whiteSpace: 'nowrap',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                        }}
                      >
                        {/* Link to open the attachment in a new tab */}
                        <Link href={a.ServerRelativeUrl} target="_blank" rel="noreferrer">
                          {a.FileName}
                        </Link>
                      </div>
                      <Text size={200}>{a.ServerRelativeUrl}</Text>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {!loadingSP && !loadError && Array.isArray(spAttachments) && spAttachments.length === 0 && (
              <Text size={200}>No existing attachments.</Text>
            )}
          </div>
        )}

        {/* ===== Hidden input (actual file control) =====
            We keep it hidden and trigger it with a styled Button for nicer UX. */}
        <input
          id={id}            /* id matches your field id prop */
          name={displayName} /* name attribute shows the displayName */
          ref={inputRef}
          type="file"
          multiple={!isSingleSelection}
          accept={accept}
          style={{ display: 'none' }}
          onChange={onFilesPicked}
          disabled={isDisabled}
        />

        {/* ===== Action buttons + helper text ===== */}
        <div className={className} style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <Button
            appearance="primary"
            icon={<AttachRegular />}
            onClick={(): void => openPicker()}
            disabled={isDisabled}
          >
            {files.length === 0
              ? isSingleSelection
                ? 'Choose file'
                : 'Choose files'
              : isSingleSelection
              ? 'Choose different file'
              : 'Add more files'}
          </Button>

          {files.length > 0 && (
            <Button
              appearance="secondary"
              onClick={(): void => clearAll()}
              icon={<DismissRegular />}
              disabled={isDisabled}
            >
              Clear
            </Button>
          )}

          {/* Describe constraints (accepted types, size, count) */}
          {(accept || isDefined(maxFileSizeMB) || (!isSingleSelection && isDefined(maxFiles))) && (
            <Text size={200} wrap>
              {accept && (
                <span>
                  Allowed: <code>{accept}</code>.{' '}
                </span>
              )}
              {isDefined(maxFileSizeMB) && <span>Max size: {maxFileSizeMB} MB/file. </span>}
              {!isSingleSelection && isDefined(maxFiles) && <span>Max files: {maxFiles}.</span>}
            </Text>
          )}
        </div>

        {/* ===== Locally selected files (not uploaded yet) ===== */}
        {files.length > 0 && (
          <div style={{ marginTop: 8, display: 'grid', gap: 6 }}>
            {files.map((f, i) => (
              <div
                key={`${f.name}-${f.size}-${i}`}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 8,
                  padding: '6px 10px',
                  borderRadius: 8,
                  border: '1px solid var(--colorNeutralStroke1)',
                }}
              >
                <DocumentRegular />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div
                    style={{
                      fontWeight: 500,
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                    }}
                  >
                    {f.name}
                  </div>
                  <Text size={200}>
                    {formatBytes(f.size)} • {f.type || 'unknown type'}
                  </Text>
                </div>

                {/* Remove a single file from the local selection */}
                <Button
                  size="small"
                  icon={<DismissRegular />}
                  onClick={handleRemove(i)}
                  disabled={isDisabled}
                  aria-label={`Remove ${f.name}`}
                />
              </div>
            ))}
          </div>
        )}

        {/* Optional helper text */}
        {description !== '' && <div className="descriptionText" style={{ marginTop: 6 }}>{description}</div>}
      </Field>
    </div>
  );
}