/**
 * Example usage:
 *
 * <FileUploadComponentSP
 *   id="attachments"
 *   displayName="Attachments"
 *   multiple
 *   accept=".pdf,.doc,.docx,image/*"
 *   maxFiles={10}
 *   maxFileSizeMB={15}
 *   isRequired={false}
 *   description="Add any supporting files."
 *   submitting={isSubmitting}
 * />
 *
 * ——————————————————————————————————————————————————————————————————————
 *
 * FileUploadComponentSP.tsx
 *
 * What this component does (high level):
 * 1) Renders a Fluent UI "file picker" that follows our global form rules.
 * 2) NEW mode (FormMode===8): show picker only; no SharePoint calls.
 * 3) EDIT/VIEW (FormMode!==8): if FormData.attachments === true, fetch existing SP attachments once
 *    and display them (read-only). If false, skip the API entirely.
 * 4) Never writes to global state on mount. Only writes on user actions.
 * 5) Parent form decides when/how to actually upload on Save.
 * 6) In multi-file mode, the button appends (Add more files). In single-file mode, it replaces (Choose different file).
 */

import * as React from 'react';
import { Field, Button, Text, useId, Link } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';
import { DynamicFormContext } from './DynamicFormContext';

// Type-only import (we read the instance from DynamicFormContext)
import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

// Project networking helper (direct import; not passed as a prop)
import { getFetchAPI } from '../Utilis/getFetchApi';

/** Public props the parent will pass in */
export interface FileUploadPropsSP {
  /** Unique key for this field in the global form data */
  id: string;

  /** Label shown above the control */
  displayName: string;

  /** Allow selecting more than one file */
  multiple?: boolean;

  /** Browser-level filter for chooser dialog (e.g., ".pdf,image/*") */
  accept?: string;

  /** Per-file max size in megabytes (omit for unlimited) */
  maxFileSizeMB?: number;

  /** Max number of files when multiple=true (omit for unlimited) */
  maxFiles?: number;

  /** If true, at least one file must be selected */
  isRequired?: boolean;

  /** Helper text under the control */
  description?: string;

  /** Optional extra class for layout */
  className?: string;

  /** When the parent is submitting, we disable the control */
  submitting?: boolean;
}

/* ------------------------------ Small helpers ------------------------------ */

const REQUIRED_MSG = 'Please select a file.';
const TOO_LARGE_MSG = (name: string, limitMB: number) =>
  `“${name}” exceeds the maximum size of ${limitMB} MB.`;
const TOO_MANY_MSG = (limit: number) =>
  `You can attach up to ${limit} file${limit === 1 ? '' : 's'}.`;

/** Narrowing helper: “defined” means not undefined */
const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

/** Safe key access on unknown shapes (we don’t re-declare the context type) */
const hasKey = (o: Record<string, unknown>, k: string) => Object.prototype.hasOwnProperty.call(o, k);
const getKey = <T,>(o: Record<string, unknown>, k: string): T => o[k] as T;

/** Read a boolean-ish flag from any of several possible keys */
const getCtxFlag = (o: Record<string, unknown>, keys: string[]): boolean =>
  keys.some(k => hasKey(o, k) && !!o[k]);

/**
 * Context may provide “lists” in different formats (array, Set, comma string, or map).
 * This function asks “is this displayName listed?” and returns a boolean either way.
 */
const isListed = (bag: unknown, name: string): boolean => {
  const needle = name.trim().toLowerCase();
  if (!bag) return false;

  if (Array.isArray(bag)) return bag.some(v => String(v).trim().toLowerCase() === needle);

  if (typeof (bag as { has?: unknown }).has === 'function') {
    for (const v of bag as Set<unknown>) if (String(v).trim().toLowerCase() === needle) return true;
    return false;
  }

  if (typeof bag === 'string') {
    return bag.split(',').map(s => s.trim().toLowerCase()).includes(needle);
  }

  if (typeof bag === 'object') {
    for (const [k, v] of Object.entries(bag as Record<string, unknown>)) {
      if (k.trim().toLowerCase() === needle && !!v) return true;
    }
  }
  return false;
};

/** Just for nicer file size display in the list */
const formatBytes = (bytes: number): string => {
  if (!Number.isFinite(bytes)) return '';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0; let n = bytes;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
  return `${n % 1 === 0 ? n.toFixed(0) : n.toFixed(2)} ${units[i]}`;
};

/** Shape of each existing SP attachment when we expand AttachmentFiles */
type SPAttachment = { FileName: string; ServerRelativeUrl: string };

/* -------------------------------- Component -------------------------------- */

export default function FileUploadComponentSP(props: FileUploadPropsSP): JSX.Element {
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
  } = props;

  /* ---- 1) Read our “global form” context (we only pull what we use) ---- */
  const formCtx = React.useContext(DynamicFormContext);
  const ctx = formCtx as unknown as Record<string, unknown>;

  const FormData = hasKey(ctx, 'FormData') ? getKey<Record<string, unknown>>(ctx, 'FormData') : undefined;
  const FormMode = hasKey(ctx, 'FormMode') ? getKey<number>(ctx, 'FormMode') : undefined;

  /** Required callbacks provided by our form provider */
  const GlobalFormData = getKey<(id: string, value: unknown) => void>(ctx, 'GlobalFormData');
  const GlobalErrorHandle = getKey<(id: string, error: string | null) => void>(ctx, 'GlobalErrorHandle');

  /** Form Customizer instance is stored on the same context (type-only import above) */
  const formCustomizerContext = hasKey(ctx, 'FormCustomizerContext')
    ? getKey<FormCustomizerContext>(ctx, 'FormCustomizerContext')
    : undefined;

  /** Mode flags: org convention is 8 = NEW, 4 = VIEW */
  const isDisplayForm = FormMode === 4; // VIEW
  const isNewMode     = FormMode === 8; // NEW

  /* ---- 2) Disabled/Hidden logic (same conventions as other fields) ---- */
  const disabledFromCtx   = getCtxFlag(ctx, ['isDisabled', 'disabled', 'formDisabled', 'Disabled']);
  const AllDisabledFields = hasKey(ctx, 'AllDisabledFields') ? ctx.AllDisabledFields : undefined;
  const AllHiddenFields   = hasKey(ctx, 'AllHiddenFields') ? ctx.AllHiddenFields : undefined;

  /* ---- 3) Local UI state ---- */
  const [required, setRequired] = React.useState<boolean>(!!isRequired);

  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || !!submitting || isListed(AllDisabledFields, displayName)
  );

  const [isHidden, setIsHidden] = React.useState<boolean>(
    isListed(AllHiddenFields, displayName)
  );

  /** Files the user picked this session (not yet uploaded anywhere) */
  const [files, setFiles] = React.useState<File[]>([]);

  /** Current validation error to show under the field (if any) */
  const [error, setError] = React.useState<string>('');

  /** Existing attachments from SharePoint (only fetched in Edit/View when needed) */
  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | null>(null);
  const [loadingSP, setLoadingSP] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');

  /** Plain unique id for the hidden <input type="file"> element */
  const inputId = useId('file');
  const inputRef = React.useRef<HTMLInputElement>(null);

  /* ---- Treat as single-file if `multiple` is false OR `maxFiles` is exactly 1 ---- */
  const isSingleSelection = (!multiple) || (maxFiles === 1);

  /* ---- 4) Keep required/disabled/hidden in sync with props + context ---- */

  React.useEffect(() => setRequired(!!isRequired), [isRequired]);

  React.useEffect(() => {
    const fromMode        = isDisplayForm;
    const fromCtx         = disabledFromCtx;
    const fromSubmitting  = !!submitting;
    const fromDisabledSet = isListed(AllDisabledFields, displayName);
    const fromHiddenSet   = isListed(AllHiddenFields, displayName);

    setIsDisabled(fromMode || fromCtx || fromDisabledSet || fromSubmitting);
    setIsHidden(fromHiddenSet);
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  /* ---- 5) EDIT/VIEW: fetch existing SharePoint attachments conditionally ---- */
  React.useEffect(() => {
    if (isNewMode) return; // NEW → explicitly no API call

    const hasAttachmentsFlag = !!(FormData && (FormData as any).attachments);
    if (!hasAttachmentsFlag) {
      // In Edit/View but item reports “no attachments”: show nothing and don’t fetch.
      setSpAttachments(null);
      setLoadError('');
      return;
    }

    // Build REST URL from the SPFx Form Customizer context
    //   _api/web/lists/getbytitle('ListTitle')/items
    //     ?$filter=Id eq <ID>
    //     &$select=AttachmentFiles
    //     &$expand=AttachmentFiles
    const listTitle = formCustomizerContext?.context?.list?.title;
    const itemId    = formCustomizerContext?.context?.item?.ID;
    if (!listTitle || !itemId) return;

    const spUrl =
      `/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items` +
      `?$filter=Id eq ${encodeURIComponent(String(itemId))}` +
      `&$select=AttachmentFiles&$expand=AttachmentFiles`;

    let cancelled = false;
    (async () => {
      setLoadingSP(true);
      setLoadError('');
      try {
        const resp = await getFetchAPI({
          spUrl,
          method: 'GET',
          headers: { Accept: 'application/json;odata=nometadata' }
        });

        // Shape is: { value: [{ AttachmentFiles: [{ FileName, ServerRelativeUrl }, ...] }] }
        const rows: any[] = resp?.value ?? [];
        const atts: SPAttachment[] = rows[0]?.AttachmentFiles ?? [];
        if (!cancelled) setSpAttachments(atts);
      } catch (e: any) {
        if (!cancelled) {
          setSpAttachments(null);
          setLoadError(e?.message || 'Failed to load attachments.');
        }
      } finally {
        if (!cancelled) setLoadingSP(false);
      }
    })();

    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isNewMode, formCustomizerContext, FormData]);

  /* ---- 6) Validation & commit helpers ---- */

  /** Check required, count, size limits (accept= is handled by the browser UI) */
  const validateSelection = React.useCallback((list: File[]): string => {
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
  }, [required, isSingleSelection, maxFiles, maxFileSizeMB]);

  /** Our single write path to the global form data (null when nothing selected) */
  const commitValue = React.useCallback((list: File[]) => {
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalFormData(id, list.length === 0 ? null : (isSingleSelection ? list[0] : list));
  }, [GlobalFormData, id, isSingleSelection]);

  /* ---- 7) User event handlers ---- */

  /** Open the hidden <input type="file"> unless disabled by context */
  const openPicker = () => { if (!isDisabled) inputRef.current?.click(); };

  /** User chose files in the dialog.
   *  - Single-file mode → replace with just the first picked file.
   *  - Multi-file mode:
   *      • On first selection: keep all picked files (capped by maxFiles if set).
   *      • On subsequent selections: append new picks up to remaining slots.
   */
  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e) => {
    const picked = Array.from(e.currentTarget.files ?? []);
    let next: File[] = [];
    let msg = '';

    if (isSingleSelection) {
      // Always keep only one file in single-selection mode
      next = picked.slice(0, 1);
    } else {
      const already = files.length;
      const capacity = isDefined(maxFiles) ? Math.max(0, maxFiles - already) : picked.length;

      if (already === 0) {
        // First selection in multi mode: take as many as allowed (could be many)
        const toTake = isDefined(maxFiles) ? Math.min(picked.length, maxFiles) : picked.length;
        next = picked.slice(0, toTake);

        if (isDefined(maxFiles) && picked.length > maxFiles) {
          msg = TOO_MANY_MSG(maxFiles);
        }
      } else {
        // Subsequent selections: append up to remaining capacity
        const toAdd = picked.slice(0, capacity);
        next = files.concat(toAdd);

        if (isDefined(maxFiles) && picked.length > capacity) {
          msg = TOO_MANY_MSG(maxFiles);
        }
      }
    }

    // If we didn't already set a "too many" message, validate normally
    if (!msg) msg = validateSelection(next);

    setFiles(next);
    setError(msg);

    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg === '' ? null : msg);
    commitValue(next);

    // Allow re-selecting the same file(s) immediately if desired
    if (inputRef.current) inputRef.current.value = '';
  };

  /** Remove a single file from the local selection */
  const removeAt = (idx: number) => {
    const next = files.filter((_, i) => i !== idx);
    const msg = validateSelection(next);

    setFiles(next);
    setError(msg);

    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg === '' ? null : msg);
    commitValue(next);
  };

  /** Clear all files from the local selection */
  const clearAll = () => {
    const msg = required ? REQUIRED_MSG : '';
    setFiles([]);
    setError(msg);

    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg || null);
    commitValue([]);

    // Let the user pick the same file again without needing to refresh
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ---- 8) Render ---- */

  if (isHidden) return <div hidden className="fieldClass" />;

  return (
    <div className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {/* Existing attachments (Edit/View only, and only if item has attachments) */}
        {!isNewMode && (
          <div style={{ marginBottom: 8 }}>
            {loadingSP && <Text size={200}>Loading attachments…</Text>}
            {!loadingSP && loadError && (
              <Text size={200} aria-live="polite">Error: {loadError}</Text>
            )}
            {!loadingSP && !loadError && spAttachments && spAttachments.length > 0 && (
              <div style={{ display: 'grid', gap: 6 }}>
                {spAttachments.map((a, i) => (
                  <div key={`${a.ServerRelativeUrl}-${i}`} style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 8,
                    padding: '6px 10px',
                    borderRadius: 8,
                    border: '1px solid var(--colorNeutralStroke1)',
                  }}>
                    <DocumentRegular />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      {/* Name is clickable and opens the SharePoint file in a new tab */}
                      <div style={{
                        fontWeight: 500,
                        whiteSpace: 'nowrap',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis'
                      }}>
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
            {!loadingSP && !loadError && spAttachments && spAttachments.length === 0 && (
              <Text size={200}>No existing attachments.</Text>
            )}
          </div>
        )}

        {/* Hidden native input that actually opens the system file picker */}
        <input
          id={inputId}
          ref={inputRef}
          type="file"
          multiple={!isSingleSelection}
          accept={accept}
          style={{ display: 'none' }}
          onChange={onFilesPicked}
          disabled={isDisabled}
        />

        {/* Trigger + Clear buttons and small “requirements” hint */}
        <div className={className} style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <Button
            appearance="primary"
            icon={<AttachRegular />}  // Fluent UI v9 paperclip icon
            onClick={openPicker}
            disabled={isDisabled}
          >
            {files.length === 0
              ? (isSingleSelection ? 'Choose file' : 'Choose files')
              : (isSingleSelection ? 'Choose different file' : 'Add more files')}
          </Button>

          {files.length > 0 && (
            <Button
              appearance="secondary"
              onClick={clearAll}
              icon={<DismissRegular />}
              disabled={isDisabled}
            >
              Clear
            </Button>
          )}

          {(accept || isDefined(maxFileSizeMB) || (!isSingleSelection && isDefined(maxFiles))) && (
            <Text size={200} wrap>
              {accept && <span>Allowed: <code>{accept}</code>. </span>}
              {isDefined(maxFileSizeMB) && <span>Max size: {maxFileSizeMB} MB/file. </span>}
              {!isSingleSelection && isDefined(maxFiles) && <span>Max files: {maxFiles}.</span>}
            </Text>
          )}
        </div>

        {/* Newly selected files (local only; these are not the historical SP attachments) */}
        {files.length > 0 && (
          <div style={{ marginTop: 8, display: 'grid', gap: 6 }}>
            {files.map((f, i) => (
              <div key={`${f.name}-${f.size}-${i}`} style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                padding: '6px 10px',
                borderRadius: 8,
                border: '1px solid var(--colorNeutralStroke1)',
              }}>
                <DocumentRegular />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{
                    fontWeight: 500,
                    whiteSpace: 'nowrap',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis'
                  }}>
                    {f.name}
                  </div>
                  <Text size={200}>{formatBytes(f.size)} • {f.type || 'unknown type'}</Text>
                </div>
                <Button
                  size="small"
                  icon={<DismissRegular />}
                  onClick={() => removeAt(i)}
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
