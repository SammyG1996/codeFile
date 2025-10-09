/**
 * FileUploadComponent.tsx
 * Fluent UI v9 file picker integrated with SPFx Form Customizer.
 * - NEW mode: pick local files; selection is written to GlobalFormData.
 * - EDIT/VIEW mode: loads existing SharePoint attachments and lets users delete
 *   them (with a confirm prompt). Deletion uses getFetchAPI (digest handled there).
 */

import * as React from 'react';
import { Field, Button, Text, Link, Spinner } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';

import { DynamicFormContext } from './DynamicFormContext';
import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { getFetchAPI } from '../Utilis/getFetchApi';

/* -------------------------------- Types -------------------------------- */

export interface FileUploadProps {
  id: string;
  displayName: string;

  multiple?: boolean;
  accept?: string;            // omit to allow any type
  maxFileSizeMB?: number;     // per-file limit
  maxFiles?: number;          // total count (only if multiple)

  isRequired?: boolean;
  description?: string;
  className?: string;
  submitting?: boolean;

  /** SPFx Form Customizer context (passed by parent). */
  context?: FormCustomizerContext;
}

type SPAttachment = { FileName: string; ServerRelativeUrl: string };

type FormCtxShape = {
  FormData?: Record<string, unknown>;
  FormMode?: number;
  GlobalFormData: (id: string, value: unknown) => void;
  // UPDATED: error uses `string | undefined` (undefined = no error)
  GlobalErrorHandle: (id: string, error: string | undefined) => void;

  isDisabled?: boolean;
  disabled?: boolean;
  formDisabled?: boolean;
  Disabled?: boolean;
  AllDisabledFields?: unknown;
  AllHiddenFields?: unknown;
};

/* ------------------------------- Helpers ------------------------------- */

const REQUIRED_MSG = 'Please select a file.';
const TOO_LARGE_MSG = (name: string, limitMB: number): string =>
  `“${name}” exceeds the maximum size of ${limitMB} MB.`;
const TOO_MANY_MSG = (limit: number): string =>
  `You can attach up to ${limit} file${limit === 1 ? '' : 's'}.`;

const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

const getCtxFlag = (o: Record<string, unknown>, keys: string[]): boolean =>
  keys.some((k) => Object.prototype.hasOwnProperty.call(o, k) && Boolean(o[k]));

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

const formatBytes = (bytes: number): string => {
  if (!Number.isFinite(bytes)) return '';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0;
  let n = bytes;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
  return `${Number.isInteger(n) ? n.toFixed(0) : n.toFixed(2)} ${units[i]}`;
};

/** Read common “has attachments” hints from the FormData snapshot. */
const readAttachmentsHint = (fd: Record<string, unknown> | undefined): boolean | undefined => {
  if (!fd) return undefined;
  const keys = ['Attachments', 'attachments', 'AttachmentCount', 'attachmentCount'] as const;
  for (const k of keys) {
    if (Object.prototype.hasOwnProperty.call(fd, k)) {
      const v = (fd as Record<string, unknown>)[k];
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v > 0;
    }
  }
  return undefined;
};

/* ------------------------------ Component ------------------------------ */

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
    context,
  } = props;

  const ctx = React.useContext(DynamicFormContext) as unknown as FormCtxShape;

  const FormData = ctx.FormData;
  const FormMode = ctx.FormMode ?? 0;
  const isNewMode = FormMode === 8;
  const isDisplayForm = FormMode === 4;

  const disabledFromCtx = getCtxFlag(ctx as unknown as Record<string, unknown>, [
    'isDisabled', 'disabled', 'formDisabled', 'Disabled',
  ]);
  const AllDisabledFields = ctx.AllDisabledFields;
  const AllHiddenFields = ctx.AllHiddenFields;

  const [required, setRequired] = React.useState<boolean>(Boolean(isRequired));
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || Boolean(submitting) || isListed(AllDisabledFields, displayName)
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(isListed(AllHiddenFields, displayName));

  const [files, setFiles] = React.useState<File[]>([]);
  const [error, setError] = React.useState<string>('');

  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | undefined>(undefined);
  const [loadingSP, setLoadingSP] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');
  const [deletingName, setDeletingName] = React.useState<string | null>(null);

  const inputRef = React.useRef<HTMLInputElement>(null);
  const isSingleSelection = !multiple || maxFiles === 1;

  React.useEffect(() => setRequired(Boolean(isRequired)), [isRequired]);

  React.useEffect(() => {
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
    if (isNewMode) return;

    const hint = readAttachmentsHint(FormData);
    if (hint === false) {
      setSpAttachments([]);
      setLoadingSP(false);
      setLoadError('');
      return;
    }

    const listTitle: string | undefined = (context as { list?: { title?: string } } | undefined)?.list?.title;
    const listGuid: string | undefined = (context as { list?: { id?: string } } | undefined)?.list?.id;
    const itemId: number | undefined = (context as { item?: { ID?: number } } | undefined)?.item?.ID;
    const baseUrl: string | undefined =
      (context as { pageContext?: { web?: { absoluteUrl?: string } } } | undefined)?.pageContext?.web?.absoluteUrl ??
      (typeof window !== 'undefined' ? window.location.origin : undefined);

    if (!baseUrl || !itemId || (!listGuid && !listTitle)) return;

    const encTitle = listTitle ? encodeURIComponent(listTitle) : '';
    const idStr = encodeURIComponent(String(itemId));

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
        setLoadingSP(false);
        setLoadError(msg);
      }
    })().catch(() => { /* noop */ });

    return (): void => { cancelled = true; };
  }, [isNewMode, FormData, context]);

  /* ----------------------- Validation & committing ---------------------- */

  const validateSelection = React.useCallback(
    (list: File[]): string => {
      if (required && list.length === 0) return REQUIRED_MSG;
      if (!isSingleSelection && isDefined(maxFiles) && list.length > maxFiles) return TOO_MANY_MSG(maxFiles);
      if (isDefined(maxFileSizeMB)) {
        const limit = maxFileSizeMB * 1024 * 1024;
        for (const f of list) if (f.size > limit) return TOO_LARGE_MSG(f.name, maxFileSizeMB);
      }
      return '';
    },
    [required, isSingleSelection, maxFiles, maxFileSizeMB]
  );

  const commitNewFiles = React.useCallback(
    (list: File[]): void => {
      const payload: unknown = list.length === 0 ? undefined : isSingleSelection ? list[0] : list;
      ctx.GlobalFormData(id, payload);
    },
    [ctx, id, isSingleSelection]
  );

  /* ------------------------------- Handlers ----------------------------- */

  const openPicker = (): void => { if (!isDisabled) inputRef.current?.click(); };

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
    // Use `undefined` to signal "no error"
    ctx.GlobalErrorHandle(id, msg === '' ? undefined : msg);
    commitNewFiles(next);

    if (inputRef.current) inputRef.current.value = '';
  };

  const removeAt = React.useCallback(
    (idx: number): void => {
      const next = files.filter((_, i) => i !== idx);
      const msg = validateSelection(next);
      setFiles(next);
      setError(msg);
      ctx.GlobalErrorHandle(id, msg === '' ? undefined : msg);
      commitNewFiles(next);
    },
    [files, validateSelection, ctx, id, commitNewFiles]
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
    ctx.GlobalErrorHandle(id, msg === '' ? undefined : msg);
    commitNewFiles([]);
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ---------------- Delete an existing SP attachment (uses getFetchAPI) ---------------- */

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
      const encFile  = encodeURIComponent(fileName);
      const idStr    = encodeURIComponent(String(itemId));

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
          // Try DELETE first (getFetchAPI adds digest)
          await getFetchAPI({
            spUrl,
            method: 'DELETE',
            headers: { 'IF-MATCH': '*', Accept: 'application/json;odata=nometadata' }
          });
          success = true;
          break;
        } catch (e1) {
          // Fallback: POST + X-HTTP-Method: DELETE
          try {
            await getFetchAPI({
              spUrl,
              method: 'POST',
              headers: {
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE',
                Accept: 'application/json;odata=nometadata'
              }
            });
            success = true;
            break;
          } catch (e2) {
            lastErr = e2 ?? e1;
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
        {/* Existing attachments (Edit/View) */}
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
                        {/* show only file name; link to open in new tab */}
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

        {/* Hidden native file input (actual picker) */}
        <input
          id={id}
          name={displayName}
          ref={inputRef}
          type="file"
          multiple={!isSingleSelection}
          accept={accept}
          style={{ display: 'none' }}
          onChange={onFilesPicked}
          disabled={isDisabled}
        />

        {/* Action row */}
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

        {/* Locally selected (new) files */}
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
                  onClick={(): void => {
                    const next = files.filter((_, idx) => idx !== i);
                    const msg = validateSelection(next);
                    setFiles(next);
                    setError(msg);
                    ctx.GlobalErrorHandle(id, msg === '' ? undefined : msg);
                    commitNewFiles(next);
                  }}
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