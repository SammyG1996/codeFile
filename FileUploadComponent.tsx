/**
 * FileUploadComponent.tsx
 * ---------------------------------------------------------------------
 * - NEW: Deleting an existing attachment now prompts for confirmation
 *   and then calls the SharePoint REST API. On success the file row
 *   disappears locally.
 * - Existing behavior:
 *   • New files are selected locally (no upload here) and committed into
 *     GlobalFormData(id) as File | File[].
 *   • Existing SP attachments are listed in Edit/View and show filename
 *     only (no path). Each row has a Delete button.
 *   • accept is optional; if you omit it, any file type is allowed.
 *   • GlobalErrorHandle gets `null` for “no error”.
 */

import * as React from 'react';
import { Field, Button, Text, Link, Spinner } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';

import { DynamicFormContext } from './DynamicFormContext';
import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { getFetchAPI } from '../Utilis/getFetchApi';

/* ============================= Types ============================= */

export interface FileUploadProps {
  id: string;                      // key used when writing to GlobalFormData
  displayName: string;             // label shown in the UI
  multiple?: boolean;              // allow multiple selection
  accept?: string;                 // OPTIONAL, omit to allow any type
  maxFileSizeMB?: number;          // per-file size limit (MB)
  maxFiles?: number;               // overall count limit (multi only)
  isRequired?: boolean;            // require at least one file
  description?: string;            // helper text
  className?: string;              // extra CSS for action row
  submitting?: boolean;            // disable while form is saving
  context?: FormCustomizerContext; // SPFx context for REST URLs
}

type FormCtxShape = {
  FormData?: Record<string, unknown>;
  FormMode?: number;
  GlobalFormData: (id: string, value: unknown) => void;
  GlobalErrorHandle: (id: string, error: string | null) => void;

  isDisabled?: boolean;
  disabled?: boolean;
  formDisabled?: boolean;
  Disabled?: boolean;
  AllDisabledFields?: unknown;
  AllHiddenFields?: unknown;
};

type SPAttachment = { FileName: string; ServerRelativeUrl: string };

/* =========================== Utilities =========================== */

const REQUIRED_MSG = 'Please select a file.';
const TOO_LARGE_MSG = (name: string, limitMB: number): string =>
  `“${name}” exceeds the maximum size of ${limitMB} MB.`;
const TOO_MANY_MSG = (limit: number): string =>
  `You can attach up to ${limit} file${limit === 1 ? '' : 's'}.`;

const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

const getCtxFlag = (o: Record<string, unknown>, keys: string[]): boolean =>
  keys.some(k => Object.prototype.hasOwnProperty.call(o, k) && Boolean(o[k]));

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

const formatBytes = (bytes: number): string => {
  if (!Number.isFinite(bytes)) return '';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0;
  let n = bytes;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
  return `${Number.isInteger(n) ? n.toFixed(0) : n.toFixed(2)} ${units[i]}`;
};

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
    accept,                      // optional → any file type if omitted
    maxFileSizeMB,
    maxFiles,
    isRequired,
    description = '',
    className,
    submitting,
    context,
  } = props;

  // Pull the form services from your provider
  const raw = React.useContext(DynamicFormContext) as unknown as FormCtxShape;

  // Mode + form data
  const FormData = raw.FormData;
  const FormMode = raw.FormMode ?? 0;
  const isDisplayForm = FormMode === 4; // VIEW
  const isNewMode = FormMode === 8;     // NEW

  // Disable/Hide calculations
  const disabledFromCtx = getCtxFlag(raw as unknown as Record<string, unknown>, [
    'isDisabled', 'disabled', 'formDisabled', 'Disabled',
  ]);
  const AllDisabledFields = raw.AllDisabledFields;
  const AllHiddenFields = raw.AllHiddenFields;

  // ------------------------ Local state ------------------------
  const [required, setRequired] = React.useState<boolean>(Boolean(isRequired));
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || Boolean(submitting) || isListed(AllDisabledFields, displayName)
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(isListed(AllHiddenFields, displayName));

  // Files newly selected (not uploaded yet)
  const [files, setFiles] = React.useState<File[]>([]);
  const [error, setError] = React.useState<string>('');

  // Existing SP attachments + loading state + errors
  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | undefined>(undefined);
  const [loadingSP, setLoadingSP] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');
  const [deletingName, setDeletingName] = React.useState<string | null>(null);

  const inputRef = React.useRef<HTMLInputElement>(null);
  const isSingleSelection = !multiple || maxFiles === 1;

  /* ------------------------- props -> state ------------------------- */

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

  /* ----------- fetch existing attachments (Edit/View only) ----------- */

  React.useEffect((): void | (() => void) => {
    if (isNewMode) return; // no attachments yet in NEW

    const attachmentsHint = readAttachmentsHint(FormData);
    if (attachmentsHint === false) {
      setSpAttachments([]);
      setLoadingSP(false);
      setLoadError('');
      return;
    }

    // Build absolute URLs from SPFx context
    const listTitle: string | undefined = (context as { list?: { title?: string } } | undefined)?.list?.title;
    const listGuid: string | undefined = (context as { list?: { id?: string } } | undefined)?.list?.id;
    const itemId: number | undefined = (context as { item?: { ID?: number } } | undefined)?.item?.ID;
    const baseUrl: string | undefined =
      (context as { pageContext?: { web?: { absoluteUrl?: string } } } | undefined)?.pageContext?.web?.absoluteUrl ??
      (typeof window !== 'undefined' ? window.location.origin : undefined);

    if (!baseUrl || !itemId || (!listGuid && !listTitle)) return;

    const encTitle = listTitle ? encodeURIComponent(listTitle) : '';
    const idStr = encodeURIComponent(String(itemId));

    // Try multiple shapes (some helpers auto-prepend /_api)
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

          // Normalise either “{ value: [{ AttachmentFiles: [...] }] }” or “{ AttachmentFiles: [...] }”
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
                  const ServerRelativeUrl = typeof o.ServerRelativeUrl === 'string' ? o.ServerRelativeUrl : undefined;
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
        setLoadError(msg);
        setLoadingSP(false);
      }
    })().catch(() => { /* satisfy lint */ });

    return (): void => { cancelled = true; };
  }, [isNewMode, FormData, context]);

  /* ----------------- Validation & committing new files ----------------- */

  const validateSelection = React.useCallback(
    (list: File[]): string => {
      if (required && list.length === 0) return REQUIRED_MSG;

      if (!isSingleSelection && isDefined(maxFiles) && list.length > maxFiles) {
        return TOO_MANY_MSG(maxFiles);
      }

      if (isDefined(maxFileSizeMB)) {
        const perFileLimitBytes = maxFileSizeMB * 1024 * 1024;
        for (const f of list) if (f.size > perFileLimitBytes) return TOO_LARGE_MSG(f.name, maxFileSizeMB);
      }
      return '';
    },
    [required, isSingleSelection, maxFiles, maxFileSizeMB]
  );

  const commitNewFiles = React.useCallback(
    (list: File[]): void => {
      const payload: unknown = list.length === 0 ? undefined : isSingleSelection ? list[0] : list;
      raw.GlobalFormData(id, payload);
    },
    [raw, id, isSingleSelection]
  );

  /* ----------------------------- Handlers ----------------------------- */

  const openPicker = (): void => {
    if (!isDisabled) inputRef.current?.click();
  };

  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e): void => {
    const picked = Array.from(e.currentTarget.files ?? []);

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

    if (inputRef.current) inputRef.current.value = '';
  };

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

  const handleRemove = React.useCallback(
    (idx: number): React.MouseEventHandler<HTMLButtonElement> =>
      (): void => removeAt(idx),
    [removeAt]
  );

  const clearAll = (): void => {
    const msg = required ? REQUIRED_MSG : '';
    setFiles([]);
    setError(msg);
    raw.GlobalErrorHandle(id, msg === '' ? null : msg);
    commitNewFiles([]);
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ----------------------- Delete existing attachment ----------------------- */
  /**
   * Confirms and then deletes an existing attachment using SharePoint REST.
   * On success we remove it from local state.
   */
  const deleteExistingAttachment = React.useCallback(
    async (fileName: string): Promise<void> => {
      if (!context) return;

      // Confirm with the user first
      const ok = window.confirm(`Are you sure you want to delete "${fileName}"?`);
      if (!ok) return;

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

      // Try a few endpoint shapes; one will fit your environment
      const urls: string[] = [];
      if (listGuid) {
        urls.push(
          `${baseUrl}/_api/web/lists(guid'${listGuid}')/items(${idStr})/AttachmentFiles('${encFile}')`,
          `${baseUrl}/web/lists(guid'${listGuid}')/items(${idStr})/AttachmentFiles('${encFile}')`
        );
      }
      if (listTitle) {
        urls.push(
          `${baseUrl}/_api/web/lists/getbytitle('${encTitle}')/items(${idStr})/AttachmentFiles('${encFile}')`,
          `${baseUrl}/web/lists/getbytitle('${encTitle}')/items(${idStr})/AttachmentFiles('${encFile}')`
        );
      }

      setDeletingName(fileName);
      let success = false;
      let lastErr: unknown = null;

      for (const spUrl of urls) {
        try {
          // Preferred: true DELETE
          await getFetchAPI({
            spUrl,
            method: 'DELETE',
            headers: { 'IF-MATCH': '*' },
          });
          success = true;
          break;
        } catch (e1: unknown) {
          lastErr = e1;
          // Fallback: POST with X-HTTP-Method override
          try {
            await getFetchAPI({
              spUrl,
              method: 'POST',
              headers: {
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE',
              },
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
        setSpAttachments(prev =>
          Array.isArray(prev) ? prev.filter(a => a.FileName !== fileName) : prev
        );
      } else {
        const msg = lastErr instanceof Error ? lastErr.message : 'Failed to delete the attachment.';
        setLoadError(msg);
      }
    },
    [context]
  );

  /* ------------------------------- Render ------------------------------- */

  if (isHidden) return <div hidden className="fieldClass" />;

  return (
    <div className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {/* ===== Existing SP attachments (Edit/View) ===== */}
        {!isNewMode && (
          <div style={{ marginBottom: 8 }}>
            {loadingSP && <Text size={200}><Spinner size="tiny" />&nbsp;Loading attachments…</Text>}

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
                        {/* Show ONLY filename (no path) */}
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
                        onClick={(): void => { void deleteExistingAttachment(a.FileName); }}
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

        {/* ===== Hidden input (actual file control) ===== */}
        <input
          id={id}
          name={displayName}
          ref={inputRef}
          type="file"
          multiple={!isSingleSelection}
          accept={accept /* omit to allow any type */}
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