/**
 * FileUploadComponentSP.tsx
 *
 * Summary
 * - File upload that understands SPFx Form Customizer modes (NEW vs EDIT/VIEW).
 * - NEW mode: render picker; no network.
 * - EDIT/VIEW mode: if FormData.attachments === true, fetch existing attachments and render them; else show blank.
 * - Validations: required, accept (MIME/ext), max file size, max files.
 * - Disabled = (FormMode===4) OR (context disabled flags) OR (AllDisabledFields) OR (submitting).
 * - Hidden  = (AllHiddenFields).
 *
 * Behavior
 * - No global writes on mount.
 * - Calls GlobalFormData only after user selects/clears files.
 * - GlobalErrorHandle fires only after first interaction (touched).
 *
 * Notes
 * - Persists selected File(s) to GlobalFormData as File | File[] (or null when empty).
 * - Existing SP attachments are displayed (name + link) but **not** auto-written to GlobalFormData.
 */

import * as React from 'react';
import { Field, Button, Text, useId, Link } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';
import { DynamicFormContext } from './DynamicFormContext';

// If you have the type, you can uncomment:
// import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

export interface FileUploadPropsSP {
  id: string;
  displayName: string;

  // Uploader options
  multiple?: boolean;
  accept?: string;
  maxFileSizeMB?: number;
  maxFiles?: number;
  isRequired?: boolean;

  // UI/help
  description?: string;
  className?: string;
  submitting?: boolean;

  // SP bits (pass these in from your form customizer)
  /** The Form Customizer context (used to build the REST URL). */
  formCustomizerContext: any; // replace with FormCustomizerContext if you prefer

  /** Your existing helper that wraps SPHttpClient/fetch. */
  getFetchAPI: (init: { spUrl: string; method?: string; headers?: Record<string, string> }) => Promise<any>;
}

/* ---------- Helpers ---------- */

const REQUIRED_MSG = 'Please select a file.';
const TOO_LARGE_MSG = (name: string, limitMB: number) =>
  `“${name}” exceeds the maximum size of ${limitMB} MB.`;
const TOO_MANY_MSG = (limit: number) =>
  `You can attach up to ${limit} file${limit === 1 ? '' : 's'}.`;

const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;
const toBool = (v: unknown): boolean => !!v;
const hasKey = (o: Record<string, unknown>, k: string) => Object.prototype.hasOwnProperty.call(o, k);
const getKey = <T,>(o: Record<string, unknown>, k: string): T => o[k] as T;
const getCtxFlag = (o: Record<string, unknown>, keys: string[]): boolean => keys.some(k => hasKey(o, k) && !!o[k]);
const isListed = (bag: unknown, name: string): boolean => {
  const needle = name.trim().toLowerCase();
  if (!bag) return false;
  if (Array.isArray(bag)) return bag.some(v => String(v).trim().toLowerCase() === needle);
  if (typeof (bag as { has?: unknown }).has === 'function') {
    for (const v of bag as Set<unknown>) if (String(v).trim().toLowerCase() === needle) return true;
    return false;
  }
  if (typeof bag === 'string') return bag.split(',').map(s => s.trim().toLowerCase()).includes(needle);
  if (typeof bag === 'object') {
    for (const [k, v] of Object.entries(bag as Record<string, unknown>))
      if (k.trim().toLowerCase() === needle && toBool(v)) return true;
  }
  return false;
};

const formatBytes = (bytes: number): string => {
  if (!Number.isFinite(bytes)) return '';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0; let n = bytes;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
  return `${n % 1 === 0 ? n.toFixed(0) : n.toFixed(2)} ${units[i]}`;
};

/* ---------- Types ---------- */

type SPAttachment = { FileName: string; ServerRelativeUrl: string };

/* ---------- Component ---------- */

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

  // Dynamic form context (same pattern as your SingleLineComponent)
  const formCtx = React.useContext(DynamicFormContext);
  const ctx = formCtx as unknown as Record<string, unknown>;

  const FormData = hasKey(ctx, 'FormData') ? getKey<Record<string, unknown>>(ctx, 'FormData') : undefined;
  const FormMode = hasKey(ctx, 'FormMode') ? getKey<number>(ctx, 'FormMode') : undefined;
  const GlobalFormData = getKey<(id: string, value: unknown) => void>(ctx, 'GlobalFormData');
  const GlobalErrorHandle = getKey<(id: string, error: string | null) => void>(ctx, 'GlobalErrorHandle');

  const isDisplayForm = FormMode === 4;     // VIEW
  const isNewMode    = FormMode === 8;     // NEW (matches your existing convention)

  const disabledFromCtx  = getCtxFlag(ctx, ['isDisabled', 'disabled', 'formDisabled', 'Disabled']);
  const AllDisabledFields = hasKey(ctx, 'AllDisabledFields') ? ctx.AllDisabledFields : undefined;
  const AllHiddenFields   = hasKey(ctx, 'AllHiddenFields') ? ctx.AllHiddenFields : undefined;

  // Local flags & state
  const [required, setRequired] = React.useState<boolean>(!!isRequired);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || !!submitting || isListed(AllDisabledFields, displayName)
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(isListed(AllHiddenFields, displayName));

  const [files, setFiles] = React.useState<File[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Existing SP attachments (EDIT/VIEW only, conditional on FormData.attachments)
  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | null>(null);
  const [loadingSP, setLoadingSP] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');

  const inputId = useId('file');
  const inputRef = React.useRef<HTMLInputElement>(null);

  /* ---------- effects ---------- */

  React.useEffect(() => setRequired(!!isRequired), [isRequired]);

  React.useEffect(() => {
    const fromMode = isDisplayForm;
    const fromCtx = disabledFromCtx;
    const fromSubmitting = !!submitting;
    const fromDisabledList = isListed(AllDisabledFields, displayName);
    const fromHiddenList = isListed(AllHiddenFields, displayName);
    setIsDisabled(fromMode || fromCtx || fromDisabledList || fromSubmitting);
    setIsHidden(fromHiddenList);
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  // === SharePoint requirement from your screenshot ===
  // Only in EDIT/VIEW and only if FormData.attachments === true do we fetch AttachmentFiles
  React.useEffect(() => {
    if (isNewMode) return; // NEW → skip
    const hasAttachmentsFlag = !!(FormData && (FormData as any).attachments);
    if (!hasAttachmentsFlag) {
      setSpAttachments(null);
      setLoadError('');
      return; // EDIT/VIEW but no attachments → blank component (no API call)
    }

    // Build REST URL from FormCustomizerContext
    // _api/web/lists/getbytitle('ListTitle')/items?$filter=Id eq <ID>&$select=AttachmentFiles&$expand=AttachmentFiles
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

        // Response shape: { value: [{ AttachmentFiles: [{ FileName, ServerRelativeUrl }, ...] }] }
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

  /* ---------- validation & commit ---------- */

  const validateSelection = React.useCallback((list: File[]): string => {
    if (required && list.length === 0) return REQUIRED_MSG;
    if (multiple && isDefined(maxFiles) && list.length > maxFiles) return TOO_MANY_MSG(maxFiles);
    if (isDefined(maxFileSizeMB)) {
      const limit = maxFileSizeMB * 1024 * 1024;
      for (const f of list) if (f.size > limit) return TOO_LARGE_MSG(f.name, maxFileSizeMB);
    }
    return '';
  }, [required, multiple, maxFiles, maxFileSizeMB]);

  const commitValue = React.useCallback((list: File[]) => {
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalFormData(id, list.length === 0 ? null : (multiple ? list : list[0]));
  }, [GlobalFormData, id, multiple]);

  /* ---------- handlers ---------- */

  const openPicker = () => { if (!isDisabled) inputRef.current?.click(); };

  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e) => {
    setTouched(true);
    const list = Array.from(e.currentTarget.files ?? []);
    const msg = validateSelection(list);
    setFiles(list);
    setError(msg);
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg === '' ? null : msg);
    commitValue(list);
  };

  const removeAt = (idx: number) => {
    setTouched(true);
    const next = files.filter((_, i) => i !== idx);
    const msg = validateSelection(next);
    setFiles(next);
    setError(msg);
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg === '' ? null : msg);
    commitValue(next);
  };

  const clearAll = () => {
    setTouched(true);
    setFiles([]);
    setError(required ? REQUIRED_MSG : '');
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, (required ? REQUIRED_MSG : '') || null);
    commitValue([]);
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ---------- render ---------- */

  // If hidden by context, bail early
  if (isHidden) return <div hidden className="fieldClass" />;

  return (
    <div className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {/* EXISTING ATTACHMENTS (EDIT/VIEW, only when FormData.attachments === true) */}
        {!isNewMode && (
          <div style={{ marginBottom: 8 }}>
            {loadingSP && <Text size={200}>Loading attachments…</Text>}
            {!loadingSP && loadError && <Text size={200} aria-live="polite">Error: {loadError}</Text>}
            {!loadingSP && !loadError && spAttachments && spAttachments.length > 0 && (
              <div style={{ display: 'grid', gap: 6 }}>
                {spAttachments.map((a, i) => (
                  <div key={`${a.ServerRelativeUrl}-${i}`} style={{
                    display: 'flex', alignItems: 'center', gap: 8,
                    padding: '6px 10px', borderRadius: 8, border: '1px solid var(--colorNeutralStroke1)',
                  }}>
                    <DocumentRegular />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 500, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
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

        {/* Hidden native input */}
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

        {/* Trigger + actions row */}
        <div className={className} style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <Button
            appearance="primary"
            icon={<AttachRegular />}   // <- replaced UploadRegular
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

        {/* Selected files list (local/new picks) */}
        {files.length > 0 && (
          <div style={{ marginTop: 8, display: 'grid', gap: 6 }}>
            {files.map((f, i) => (
              <div key={`${f.name}-${f.size}-${i}`} style={{
                display: 'flex', alignItems: 'center', gap: 8,
                padding: '6px 10px', borderRadius: 8, border: '1px solid var(--colorNeutralStroke1)',
              }}>
                <DocumentRegular />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 500, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
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
