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
 *   formCustomizerContext={FormCustomizerContext}   // pass your SPFx Form Customizer context
 *   getFetchAPI={getFetchAPI}                       // pass your SPHttpClient/fetch wrapper
 * />
 *
 * ——————————————————————————————————————————————————————————————————————
 *
 * FileUploadComponentSP.tsx
 *
 * What this component does (high level):
 * 1) Renders a Fluent UI "file picker" field that follows the same rules as our other fields
 *    (disabled/hidden lists, required validation, etc.).
 * 2) NEW mode (FormMode===8): just show the file picker. We do NOT call SharePoint in NEW mode.
 * 3) EDIT/VIEW mode (FormMode!==8): if the current item reports `FormData.attachments === true`,
 *    we call SharePoint once to fetch and display existing attachments; otherwise we show nothing.
 * 4) We NEVER write to global state on mount. We only call:
 *       - GlobalFormData(id, value) when the user selects/clears files
 *       - GlobalErrorHandle(id, message) when we need to surface a validation error
 * 5) The parent form can then decide what to do with the files (e.g., upload on Save).
 */

import * as React from 'react';
import { Field, Button, Text, useId, Link } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';
import { DynamicFormContext } from './DynamicFormContext';

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

  /** SPFx Form Customizer context (we use it only to build the REST URL in Edit/View) */
  formCustomizerContext: any;

  /** Your wrapper around SPHttpClient/fetch; must support GET to SharePoint REST */
  getFetchAPI: (init: { spUrl: string; method?: string; headers?: Record<string, string> }) => Promise<any>;
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
    formCustomizerContext,
    getFetchAPI,
  } = props;

  /* ---- 1) Read our “global form” context, but don’t assume a strict shape ----
     We only pull the pieces we actually use: FormMode, FormData, and the 2 callbacks. */
  const formCtx = React.useContext(DynamicFormContext);
  const ctx = formCtx as unknown as Record<string, unknown>;

  const FormData = hasKey(ctx, 'FormData') ? getKey<Record<string, unknown>>(ctx, 'FormData') : undefined;
  const FormMode = hasKey(ctx, 'FormMode') ? getKey<number>(ctx, 'FormMode') : undefined;

  /** These two are required to exist on the provider (we “assert-read” them). */
  const GlobalFormData = getKey<(id: string, value: unknown) => void>(ctx, 'GlobalFormData');
  const GlobalErrorHandle = getKey<(id: string, error: string | null) => void>(ctx, 'GlobalErrorHandle');

  /** Mode flags: our org uses 8 = NEW, 4 = VIEW */
  const isDisplayForm = FormMode === 4; // VIEW
  const isNewMode     = FormMode === 8; // NEW

  /* ---- 2) Disabled/Hidden calculation mirrors your other fields ---- */
  const disabledFromCtx   = getCtxFlag(ctx, ['isDisabled', 'disabled', 'formDisabled', 'Disabled']);
  const AllDisabledFields = hasKey(ctx, 'AllDisabledFields') ? ctx.AllDisabledFields : undefined;
  const AllHiddenFields   = hasKey(ctx, 'AllHiddenFields') ? ctx.AllHiddenFields : undefined;

  /* ---- 3) Local UI state: required, disabled/hidden, selected files, errors ---- */
  const [required, setRequired] = React.useState<boolean>(!!isRequired);

  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm                                  // view mode
    || disabledFromCtx                              // global disabled flag
    || !!submitting                                 // parent is submitting
    || isListed(AllDisabledFields, displayName)     // in the disabled list
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

  /* ---- 4) Keep local required/disabled/hidden in sync with props + context ---- */

  // If the parent toggles isRequired at runtime, reflect it locally.
  React.useEffect(() => setRequired(!!isRequired), [isRequired]);

  // Recompute disabled/hidden when any of the inputs that can influence them change.
  React.useEffect(() => {
    const fromMode        = isDisplayForm;
    const fromCtx         = disabledFromCtx;
    const fromSubmitting  = !!submitting;
    const fromDisabledSet = isListed(AllDisabledFields, displayName);
    const fromHiddenSet   = isListed(AllHiddenFields, displayName);

    setIsDisabled(fromMode || fromCtx || fromDisabledSet || fromSubmitting);
    setIsHidden(fromHiddenSet);
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  /* ---- 5) EDIT/VIEW requirement: conditionally fetch existing SharePoint attachments ----
     Per the requirement, we only call SharePoint in EDIT/VIEW *if* the current item has
     `attachments === true` in FormData. Otherwise we don’t call the API at all. */
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

  /* ---- 6) Validation & commit helpers ----
     We validate locally for UX, then notify the parent via GlobalErrorHandle.
     We commit the selected file(s) to GlobalFormData so the parent can save/upload later. */

  /** Check required, count, size limits (accept= is handled by the browser UI) */
  const validateSelection = React.useCallback((list: File[]): string => {
    if (required && list.length === 0) return REQUIRED_MSG;

    if (multiple && isDefined(maxFiles) && list.length > maxFiles) {
      return TOO_MANY_MSG(maxFiles);
    }

    if (isDefined(maxFileSizeMB)) {
      const perFileLimitBytes = maxFileSizeMB * 1024 * 1024;
      for (const f of list) {
        if (f.size > perFileLimitBytes) return TOO_LARGE_MSG(f.name, maxFileSizeMB);
      }
    }
    return '';
  }, [required, multiple, maxFiles, maxFileSizeMB]);

  /** Our single write path to the global form data (null when nothing selected) */
  const commitValue = React.useCallback((list: File[]) => {
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalFormData(id, list.length === 0 ? null : (multiple ? list : list[0]));
  }, [GlobalFormData, id, multiple]);

  /* ---- 7) User event handlers (only places we write to globals) ---- */

  /** Open the hidden <input type="file"> unless disabled by context */
  const openPicker = () => { if (!isDisabled) inputRef.current?.click(); };

  /** User chose files in the dialog */
  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e) => {
    const list = Array.from(e.currentTarget.files ?? []);
    const msg = validateSelection(list);

    setFiles(list);
    setError(msg);

    // Announce the current error state to the parent (null clears an existing error)
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg === '' ? null : msg);

    // Commit the selection so the parent has the files (or null)
    commitValue(list);
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
    setFiles([]);
    setError(required ? REQUIRED_MSG : '');

    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, (required ? REQUIRED_MSG : '') || null);
    commitValue([]);

    // Let the user pick the same file again without needing to refresh
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ---- 8) Render ----
     - We honor “hidden” by returning a hidden wrapper (keeps layout parity with other fields).
     - In Edit/View with attachments=true, we show existing SP attachments above the picker.
     - Then we show the trigger/clear buttons and a list of newly selected (local) files.
   */

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
          multiple={multiple}
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
            {files.length ? 'Choose different file' : (multiple ? 'Choose files' : 'Choose file')}
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

          {(accept || maxFileSizeMB || (multiple && maxFiles)) && (
            <Text size={200} wrap>
              {accept && <span>Allowed: <code>{accept}</code>. </span>}
              {isDefined(maxFileSizeMB) && <span>Max size: {maxFileSizeMB} MB/file. </span>}
              {multiple && isDefined(maxFiles) && <span>Max files: {maxFiles}.</span>}
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
