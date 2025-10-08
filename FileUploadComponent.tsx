/**
 * FileUploadComponent.tsx
 * -----------------------------------------------------------------------------
 * Purpose
 * - Fluent UI v9 file picker with:
 *   • NEW files: selected locally (no upload here) and written to GlobalFormData.
 *   • EXISTING SharePoint attachments (Edit/View): listed with a Delete button.
 *     On delete we show a confirmation and then call the SharePoint REST API.
 *     If the call succeeds, the row is removed locally.
 *
 * What this component DOES NOT do
 * - It does not upload files to SharePoint. Your submitter component will read
 *   GlobalFormData(id) and perform the actual upload when the form is saved.
 *
 * Form integration (DynamicFormContext contract)
 * - GlobalFormData(id, value): we write either:
 *     • undefined  → no files selected
 *     • File       → single selection
 *     • File[]     → multiple selection
 * - GlobalErrorHandle(id, error): we report a string on error, or null for “no error”.
 *
 * Requirements that are implemented here
 * - Label equals `displayName`. Input’s id equals `id` (from props).
 * - Optional `accept` (omit to allow any type).
 * - Optional `maxFileSizeMB` and `maxFiles`.
 * - Button text switches between “Choose file(s)”, “Choose different file”, “Add more files”.
 * - Show filename only (no path) for existing attachments.
 * - Existing attachment deletion uses `window.confirm` and SharePoint REST.
 * - Tidy, junior-friendly comments throughout.
 */

import * as React from 'react';
import { Field, Button, Text, Link, Spinner } from '@fluentui/react-components';
import {
  DismissRegular,
  DocumentRegular,
  AttachRegular,
} from '@fluentui/react-icons';

import { DynamicFormContext } from './DynamicFormContext';
import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { getFetchAPI } from '../Utilis/getFetchApi';

/* ------------------------------- Types -------------------------------- */

export interface FileUploadProps {
  /** Key used when writing to GlobalFormData and GlobalErrorHandle */
  id: string;
  /** Field label shown to the user */
  displayName: string;

  /** Allow selecting multiple files (default: false) */
  multiple?: boolean;

  /** OPTIONAL: file type filter, e.g. ".pdf,.docx,image/*". Omit to accept any. */
  accept?: string;

  /** OPTIONAL: per-file size limit in MB (e.g., 15) */
  maxFileSizeMB?: number;

  /** OPTIONAL: maximum count for multi-select (ignored for single-select) */
  maxFiles?: number;

  /** Mark this field as required */
  isRequired?: boolean;

  /** Helper text under the picker */
  description?: string;

  /** Optional class for the action row */
  className?: string;

  /** Disable interactions while the form is saving */
  submitting?: boolean;

  /**
   * SPFx Form Customizer context (we read site/list/item and absolute url).
   * This is passed in as a prop from the parent, per your project’s pattern.
   */
  context?: FormCustomizerContext;
}

/** What a SharePoint Attachment looks like in the responses we normalize. */
type SPAttachment = { FileName: string; ServerRelativeUrl: string };

/** Shape we expect from the DynamicFormContext provider. */
type FormCtxShape = {
  FormData?: Record<string, unknown>;
  FormMode?: number;
  GlobalFormData: (id: string, value: unknown) => void;
  GlobalErrorHandle: (id: string, error: string | null) => void;

  // Optional convenience flags/collections that may exist in your provider
  isDisabled?: boolean;
  disabled?: boolean;
  formDisabled?: boolean;
  Disabled?: boolean;
  AllDisabledFields?: unknown;
  AllHiddenFields?: unknown;
};

/* ------------------------------ Helpers ------------------------------- */

/** Standard messages */
const REQUIRED_MSG = 'Please select a file.';
const TOO_LARGE_MSG = (name: string, limitMB: number): string =>
  `“${name}” exceeds the maximum size of ${limitMB} MB.`;
const TOO_MANY_MSG = (limit: number): string =>
  `You can attach up to ${limit} file${limit === 1 ? '' : 's'}.`;

/** Type guard for defined values */
const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

/** OR over a set of boolean-ish flags on an object */
const getCtxFlag = (o: Record<string, unknown>, keys: string[]): boolean =>
  keys.some((k) => Object.prototype.hasOwnProperty.call(o, k) && Boolean(o[k]));

/** Membership helper that works with array/set/string/object maps */
const isListed = (bag: unknown, name: string): boolean => {
  const needle = name.trim().toLowerCase();
  if (bag === null || bag === undefined) return false;

  if (Array.isArray(bag)) return bag.some((v) => String(v).trim().toLowerCase() === needle);

  if (typeof (bag as { has?: unknown }).has === 'function') {
    for (const v of bag as Set<unknown>) if (String(v).trim().toLowerCase() === needle) return true;
    return false;
  }

  if (typeof bag === 'string') return bag.split(',').map((s) => s.trim().toLowerCase()).includes(needle);

  if (typeof bag === 'object') {
    for (const [k, v] of Object.entries(bag as Record<string, unknown>)) {
      if (k.trim().toLowerCase() === needle && Boolean(v)) return true;
    }
  }
  return false;
};

/** Pretty-print a file size */
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

/**
 * Try to infer whether the current item has attachments from the FormData
 * snapshot your provider gave us. Different sources use different keys.
 */
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

/* ----------------------------- Component ------------------------------ */

export default function FileUploadComponent(props: FileUploadProps): JSX.Element {
  const {
    id,
    displayName,
    multiple = false,
    accept, // optional → any file type if omitted
    maxFileSizeMB,
    maxFiles,
    isRequired,
    description = '',
    className,
    submitting,
    context,
  } = props;

  // Pull core services from your DynamicFormContext provider
  const raw = React.useContext(DynamicFormContext) as unknown as FormCtxShape;

  // Mode + form data (we rely on your provider’s numbers: 8=NEW, 4=DISPLAY)
  const FormData = raw.FormData;
  const FormMode = raw.FormMode ?? 0;
  const isDisplayForm = FormMode === 4; // VIEW
  const isNewMode = FormMode === 8; // NEW

  // Disabled/hidden logic from context
  const disabledFromCtx = getCtxFlag(raw as unknown as Record<string, unknown>, [
    'isDisabled',
    'disabled',
    'formDisabled',
    'Disabled',
  ]);
  const AllDisabledFields = raw.AllDisabledFields;
  const AllHiddenFields = raw.AllHiddenFields;

  // ---------------------------- Local state ----------------------------

  // Required & disabled reflect props/context/mode
  const [required, setRequired] = React.useState<boolean>(Boolean(isRequired));
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || Boolean(submitting) || isListed(AllDisabledFields, displayName)
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(isListed(AllHiddenFields, displayName));

  // New (local-only) selected files
  const [files, setFiles] = React.useState<File[]>([]);
  const [error, setError] = React.useState<string>('');

  // Existing SP attachments + state while loading/deleting
  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | undefined>(undefined);
  const [loadingSP, setLoadingSP] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');
  const [deletingName, setDeletingName] = React.useState<string | null>(null);

  // DOM refs
  const inputRef = React.useRef<HTMLInputElement>(null);

  // If multiple is false OR maxFiles===1, we treat as single-selection
  const isSingleSelection = !multiple || maxFiles === 1;

  /* ---------- keep state in sync with prop/context changes ----------- */

  React.useEffect((): void => {
    setRequired(Boolean(isRequired));
  }, [isRequired]);

  React.useEffect((): void => {
    const fromMode = isDisplayForm;
    const fromCtx = disabledFromCtx;
    const fromSubmitting = Boolean(submitting);
    const fromDisabledList = isListed(AllDisabledFields, displayName);
    const fromHiddenList = isListed(AllHiddenFields, displayName);

    setIsDisabled(fromMode || fromCtx || fromDisabledList || fromSubmitting);
    setIsHidden(fromHiddenList);
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  /* ------------------ Load existing attachments (Edit/View) ------------------ */

  React.useEffect((): void | (() => void) => {
    // In NEW mode there can't be existing attachments
    if (isNewMode) return;

    // If provider told us there are no attachments, skip any network call
    const attachmentsHint = readAttachmentsHint(FormData);
    if (attachmentsHint === false) {
      setSpAttachments([]);
      setLoadingSP(false);
      setLoadError('');
      return;
    }

    // Build the REST URLs from the SPFx context
    const listTitle: string | undefined = (context as { list?: { title?: string } } | undefined)?.list?.title;
    const listGuid: string | undefined = (context as { list?: { id?: string } } | undefined)?.list?.id;
    const itemId: number | undefined = (context as { item?: { ID?: number } } | undefined)?.item?.ID;
    const baseUrl: string | undefined =
      (context as { pageContext?: { web?: { absoluteUrl?: string } } } | undefined)?.pageContext?.web?.absoluteUrl ??
      (typeof window !== 'undefined' ? window.location.origin : undefined);

    // If we can't assemble a valid request, stop quietly
    if (!baseUrl || !itemId || (!listGuid && !listTitle)) return;

    const encTitle = listTitle ? encodeURIComponent(listTitle) : '';
    const idStr = encodeURIComponent(String(itemId));

    // Try multiple shapes (some libs prepend /_api for us, some don’t)
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

    (async (): Promise<void> => {
      setLoadingSP(true);
      setLoadError('');

      let success = false;
      let lastErr: unknown = null;

      for (const spUrl of urls) {
        if (cancelled) return;
        try {
          const respUnknown: unknown = await getFetchAPI({
            spUrl,
            method: 'GET',
            headers: { Accept: 'application/json;odata=nometadata' },
          });

          // Normalize either shape:
          //  a) { value: [{ AttachmentFiles: [...] }] }
          //  b) { AttachmentFiles: [...] }
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

          const atts: SPAttachment[] = Array.isArray(attsRaw)
            ? (attsRaw as unknown[]).map((x): SPAttachment | undefined => {
                if (x && typeof x === 'object') {
                  const o = x as Record<string, unknown>;
                  const FileName = typeof o.FileName === 'string' ? o.FileName : undefined;
                  const ServerRelativeUrl =
                    typeof o.ServerRelativeUrl === 'string' ? o.ServerRelativeUrl : undefined;
                  if (FileName && ServerRelativeUrl) return { FileName, ServerRelativeUrl };
                }
                return undefined;
              }).filter((x): x is SPAttachment => !!x)
            : [];

          setSpAttachments(atts);
          setLoadingSP(false);
          success = true;
          break;
        } catch (e: unknown) {
          lastErr = e;
        }
      }

      if (!success && !cancelled) {
        const msg = lastErr instanceof Error ? lastErr.message : 'Failed to load attachments.';
        setSpAttachments(undefined);
        setLoadingSP(false);
        setLoadError(msg);
      }
    })().catch(() => {
      /* swallow to satisfy lint */
    });

    return (): void => {
      cancelled = true;
    };
  }, [isNewMode, FormData, context]);

  /* ------------------- Validation & committing new files ------------------- */

  /** Validate the local selection and return an error message or "" */
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

  /** Write current selection into GlobalFormData in the agreed shape */
  const commitNewFiles = React.useCallback(
    (list: File[]): void => {
      const payload: unknown = list.length === 0 ? undefined : isSingleSelection ? list[0] : list;
      raw.GlobalFormData(id, payload);
    },
    [raw, id, isSingleSelection]
  );

  /* ------------------------------- Handlers ------------------------------- */

  /** Programmatically open the hidden input */
  const openPicker = (): void => {
    if (!isDisabled) inputRef.current?.click();
  };

  /** Handle picking files from disk */
  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e): void => {
    const picked = Array.from(e.currentTarget.files ?? []);

    // Merge with existing selection when multi; replace when single
    let next: File[] = [];
    let msg = '';

    if (isSingleSelection) {
      next = picked.slice(0, 1);
    } else {
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

    if (!msg) msg = validateSelection(next);

    setFiles(next);
    setError(msg);
    raw.GlobalErrorHandle(id, msg === '' ? null : msg); // “no error” → null
    commitNewFiles(next);

    // Clear the native input so the same file can be re-picked later
    if (inputRef.current) inputRef.current.value = '';
  };

  /** Remove a locally selected file by index */
  const removeAt = React.useCallback(
    (idx: number): void => {
      const next = files.filter((_, i) => i !== idx);
      const msg = validateSelection(next);

      setFiles(next);
      setError(msg);
      raw.GlobalErrorHandle(id, msg === '' ? null : msg);
      commitNewFiles(next);
    },
    [files, validateSelection, raw, id, commitNewFiles]
  );

  /** Build a click handler that removes a specific index */
  const handleRemove = React.useCallback(
    (idx: number): React.MouseEventHandler<HTMLButtonElement> =>
      (): void => removeAt(idx),
    [removeAt]
  );

  /** Clear all local files */
  const clearAll = (): void => {
    const msg = required ? REQUIRED_MSG : '';
    setFiles([]);
    setError(msg);
    raw.GlobalErrorHandle(id, msg === '' ? null : msg);
    commitNewFiles([]);
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ---------------- Delete an existing SP attachment (with confirm) ---------------- */

  /**
   * Ask for confirmation, then call SharePoint REST to delete an attachment.
   * Notes:
   * - We try a few URL shapes (with/without _api, by GUID or list title).
   * - We pass `expectJson: false` because DELETE often returns 204 No Content.
   * - We pass `includeDigest: true` so the helper adds a request digest.
   */
  const deleteExistingAttachment = React.useCallback(
    async (fileName: string): Promise<void> => {
      if (!context) return;

      if (!window.confirm(`Are you sure you want to delete "${fileName}"?`)) return;

      const listTitle: string | undefined = (context as { list?: { title?: string } } | undefined)?.list?.title;
      const listGuid: string | undefined = (context as { list?: { id?: string } } | undefined)?.list?.id;
      const itemId: number | undefined = (context as { item?: { ID?: number } } | undefined)?.item?.ID;
      const baseUrl: string | undefined =
        (context as { pageContext?: { web?: { absoluteUrl?: string } } } | undefined)?.pageContext?.web?.absoluteUrl ??
        (typeof window !== 'undefined' ? window.location.origin : undefined);

      if (!baseUrl || !itemId || (!listGuid && !listTitle)) return;

      const encTitle = listTitle ? encodeURIComponent(listTitle) : '';
      const encFile = encodeURIComponent(fileName);
      const idStr = encodeURIComponent(String(itemId));

      // Using getByFileName is clear and reliable; keep a couple variants
      const urls: string[] = [];
      if (listGuid) {
        urls.push(
          `${baseUrl}/_api/web/lists(guid'${listGuid}')/items(${idStr})/AttachmentFiles/getByFileName('${encFile}')`,
          `${baseUrl}/web/lists(guid'${listGuid}')/items(${idStr})/AttachmentFiles/getByFileName('${encFile}')`
        );
      }
      if (listTitle) {
        urls.push(
          `${baseUrl}/_api/web/lists/getbytitle('${encTitle}')/items(${idStr})/AttachmentFiles/getByFileName('${encFile}')`,
          `${baseUrl}/web/lists/getbytitle('${encTitle}')/items(${idStr})/AttachmentFiles/getByFileName('${encFile}')`
        );
      }

      setDeletingName(fileName);

      let success = false;
      let lastErr: unknown = null;

      for (const spUrl of urls) {
        try {
          // Preferred: true DELETE (often returns 204)
          await getFetchAPI({
            spUrl,
            method: 'DELETE',
            headers: { 'IF-MATCH': '*', Accept: 'application/json;odata=nometadata' },
            expectJson: false, // DELETE → no JSON body
            includeDigest: true,
            spfxContext: context,
          });
          success = true;
          break;
        } catch (e1: unknown) {
          lastErr = e1;

          // Fallback: POST override
          try {
            await getFetchAPI({
              spUrl,
              method: 'POST',
              headers: {
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE',
                Accept: 'application/json;odata=nometadata',
              },
              expectJson: false,
              includeDigest: true,
              spfxContext: context,
            });
            success = true;
            break;
          } catch (e2: unknown) {
            lastErr = e2;
          }
        }
      }

      setDeletingName(null);

      if (success) {
        // Remove from UI
        setSpAttachments((prev) => (Array.isArray(prev) ? prev.filter((a) => a.FileName !== fileName) : prev));
      } else {
        const msg = lastErr instanceof Error ? lastErr.message : 'Failed to delete the attachment.';
        setLoadError(msg);
      }
    },
    [context]
  );

  /* -------------------------------- Render ------------------------------- */

  if (isHidden) return <div hidden className="fieldClass" />;

  return (
    <div className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {/* ===== Existing SharePoint attachments (Edit/View) ===== */}
        {!isNewMode && (
          <div style={{ marginBottom: 8 }}>
            {loadingSP && (
              <Text size={200}>
                <Spinner size="tiny" />&nbsp;Loading attachments…
              </Text>
            )}

            {!loadingSP && loadError && (
              <Text size={200} aria-live="polite">
                Error: {loadError}
              </Text>
            )}

            {!loadingSP && !loadError && Array.isArray(spAttachments) && spAttachments.length > 0 && (
              <div style={{ display: 'grid', gap: 6 }}>
                {spAttachments.map((a) => {
                  const busy = deletingName === a.FileName;
                  return (
                    <div
                      key={a.ServerRelativeUrl}
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
                        {/* Show ONLY the filename (no path). The link opens the file. */}
                        <div
                          style={{
                            fontWeight: 500,
                            whiteSpace: 'nowrap',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                          }}
                          title={a.FileName}
                        >
                          <Link href={a.ServerRelativeUrl} target="_blank" rel="noreferrer">
                            {a.FileName}
                          </Link>
                        </div>
                      </div>

                      {/* Confirm + delete (makes REST call on confirm) */}
                      <Button
                        size="small"
                        appearance="secondary"
                        icon={<DismissRegular />}
                        disabled={busy || isDisabled}
                        aria-label={`Delete ${a.FileName}`}
                        onClick={(): void => {
                          void deleteExistingAttachment(a.FileName);
                        }}
                      >
                        {busy ? 'Deleting…' : 'Delete'}
                      </Button>
                    </div>
                  );
                })}
              </div>
            )}

            {!loadingSP && !loadError && Array.isArray(spAttachments) && spAttachments.length === 0 && (
              <Text size={200}>No existing attachments.</Text>
            )}
          </div>
        )}

        {/* ===== Hidden native input (actual file control) ===== */}
        <input
          id={id}                 /* per your request: use the prop id */
          name={displayName}      /* and use displayName for the name attribute */
          ref={inputRef}
          type="file"
          multiple={!isSingleSelection}
          accept={accept /* omit to allow ANY type */}
          style={{ display: 'none' }}
          onChange={onFilesPicked}
          disabled={isDisabled}
        />

        {/* ===== Action buttons + constraints text ===== */}
        <div className={className} style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <Button appearance="primary" icon={<AttachRegular />} onClick={openPicker} disabled={isDisabled}>
            {files.length === 0
              ? isSingleSelection
                ? 'Choose file'
                : 'Choose files'
              : isSingleSelection
              ? 'Choose different file'
              : 'Add more files'}
          </Button>

          {files.length > 0 && (
            <Button appearance="secondary" onClick={clearAll} icon={<DismissRegular />} disabled={isDisabled}>
              Clear
            </Button>
          )}

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

        {/* ===== Locally selected files (to be uploaded by your submitter) ===== */}
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
                    title={f.name}
                  >
                    {f.name}
                  </div>
                  <Text size={200}>
                    {formatBytes(f.size)} • {f.type || 'unknown type'}
                  </Text>
                </div>

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

        {description !== '' && <div className="descriptionText" style={{ marginTop: 6 }}>{description}</div>}
      </Field>
    </div>
  );
}