/**
 * Example usage (debugging ON by default):
 *
 * <FileUploadComponent
 *   id="attachments"
 *   displayName="Attachments"
 *   multiple
 *   accept=".pdf,.doc,.docx,image/*"
 *   maxFiles={10}
 *   maxFileSizeMB={15}
 *   isRequired={false}
 *   description="Add any supporting files."
 *   submitting={isSubmitting}
 *   // debug={false}   // turn logs off later
 * />
 */

import * as React from 'react';
import { Field, Button, Text, useId, Link } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';
import { DynamicFormContext } from './DynamicFormContext';

import type { FormCustomizerContext as SPFxFormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { getFetchAPI } from '../Utilis/getFetchApi';

/* ------------------------------ Types ------------------------------ */

export interface FileUploadProps {
  id: string;
  displayName: string;
  multiple?: boolean;
  accept?: string;
  maxFileSizeMB?: number;
  maxFiles?: number;
  isRequired?: boolean;
  description?: string;
  className?: string;
  submitting?: boolean;
  /** Turn console logging on/off (defaults to true for diagnostics). */
  debug?: boolean;
}

/** Minimal view of our form context. All fields optional on purpose. */
type FormCtxShape = {
  FormData?: Record<string, unknown>;
  FormMode?: number;
  GlobalFormData?: (id: string, value: unknown) => void;
  GlobalErrorHandle?: (id: string, error: string | undefined) => void;

  // SPFx form context instance (sometimes provided as `.context`, sometimes direct)
  FormCustomizerContext?: unknown;

  // optional flags/lists used by our field pattern
  isDisabled?: boolean;
  disabled?: boolean;
  formDisabled?: boolean;
  Disabled?: boolean;
  AllDisabledFields?: unknown;
  AllHiddenFields?: unknown;
};

type SPAttachment = { FileName: string; ServerRelativeUrl: string };

/* ------------------------------ Helpers ------------------------------ */

const REQUIRED_MSG = 'Please select a file.';
const TOO_LARGE_MSG = (name: string, limitMB: number): string =>
  `“${name}” exceeds the maximum size of ${limitMB} MB.`;
const TOO_MANY_MSG = (limit: number): string =>
  `You can attach up to ${limit} file${limit === 1 ? '' : 's'}.`;

const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

const hasKey = (o: Record<string, unknown>, k: string): boolean =>
  Object.prototype.hasOwnProperty.call(o, k);

const getCtxFlag = (o: Record<string, unknown>, keys: string[]): boolean =>
  keys.some(k => hasKey(o, k) && Boolean(o[k]));

/** Accepts array, Set, comma-string, or object-map and checks membership of displayName */
const isListed = (bag: unknown, name: string): boolean => {
  const needle = name.trim().toLowerCase();
  if (bag === null || bag === undefined) return false;

  if (Array.isArray(bag)) return bag.some(v => String(v).trim().toLowerCase() === needle);

  if (typeof (bag as { has?: unknown }).has === 'function') {
    for (const v of bag as Set<unknown>) {
      if (String(v).trim().toLowerCase() === needle) return true;
    }
    return false;
  }

  if (typeof bag === 'string') {
    return bag.split(',').map(s => s.trim().toLowerCase()).includes(needle);
  }

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
  while (n >= 1024 && i < units.length - 1) {
    n /= 1024;
    i++;
  }
  return `${Number.isInteger(n) ? n.toFixed(0) : n.toFixed(2)} ${units[i]}`;
};

/** Defensive read of a boolean-like field from an unknown object */
const readBool = (obj: unknown, key: string): boolean => {
  if (obj && typeof obj === 'object' && hasKey(obj as Record<string, unknown>, key)) {
    return Boolean((obj as Record<string, unknown>)[key]);
  }
  return false;
};

/**
 * Safely read list title & item ID from the provided SPFx Form Customizer context.
 * Supports both shapes:
 *   1) context itself: { list: { title }, item: { ID } }
 *   2) wrapped:       { context: { list: { title }, item: { ID } } }
 */
const getListTitleAndItemId = (
  ctx: unknown
): { listTitle?: string; itemId?: number; shape: 'direct' | 'wrapped' | 'unknown' } => {
  if (ctx && typeof ctx === 'object' && hasKey(ctx as Record<string, unknown>, 'context')) {
    const root = (ctx as { context: unknown }).context;
    if (!root || typeof root !== 'object') return { shape: 'wrapped' };
    const listTitle =
      hasKey(root as Record<string, unknown>, 'list') &&
      typeof (root as { list?: { title?: unknown } }).list?.title === 'string'
        ? ((root as { list?: { title?: string } }).list!.title as string)
        : undefined;

    const itemId =
      hasKey(root as Record<string, unknown>, 'item') &&
      typeof (root as { item?: { ID?: unknown } }).item?.ID === 'number'
        ? ((root as { item?: { ID?: number } }).item!.ID as number)
        : undefined;

    return { listTitle, itemId, shape: 'wrapped' };
  }

  if (ctx && typeof ctx === 'object') {
    const listTitle =
      hasKey(ctx as Record<string, unknown>, 'list') &&
      typeof (ctx as { list?: { title?: unknown } }).list?.title === 'string'
        ? ((ctx as { list?: { title?: string } }).list!.title as string)
        : undefined;

    const itemId =
      hasKey(ctx as Record<string, unknown>, 'item') &&
      typeof (ctx as { item?: { ID?: unknown } }).item?.ID === 'number'
        ? ((ctx as { item?: { ID?: number } }).item!.ID as number)
        : undefined;

    return { listTitle, itemId, shape: 'direct' };
  }

  return { shape: 'unknown' };
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
    debug = true,
  } = props;

  const log = (...args: unknown[]): void => {
    if (!debug) return;
    // eslint-disable-next-line no-console
    console.log('%c[%cFileUpload%c]', 'color:#888', 'color:#0b6', 'color:#888', ...args);
  };

  const logAtt = (...args: unknown[]): void => {
    if (!debug) return;
    // eslint-disable-next-line no-console
    console.log('%c[%cFileUpload%c] %cATTACHMENTS', 'color:#888', 'color:#0b6', 'color:#888', 'color:#06c;font-weight:bold', ...args);
  };

  log('mount: props =', { id, displayName, multiple, accept, maxFileSizeMB, maxFiles, isRequired, submitting });

  // Context (typed to our minimal shape)
  const raw = React.useContext(DynamicFormContext) as unknown as FormCtxShape;
  log('context snapshot:', raw);

  const FormData = raw.FormData;
  const FormMode = raw.FormMode;
  const GlobalFormData = raw.GlobalFormData as (id: string, value: unknown) => void;
  const GlobalErrorHandle = raw.GlobalErrorHandle as (id: string, error: string | undefined) => void;

  // Treat this as unknown, we’ll safely read the fields we need
  const formCustomizerContext: unknown = raw.FormCustomizerContext as unknown as SPFxFormCustomizerContext | unknown;

  const isDisplayForm = FormMode === 4; // VIEW
  const isNewMode = FormMode === 8; // NEW
  log('modes => { FormMode, isDisplayForm, isNewMode }', { FormMode, isDisplayForm, isNewMode });

  const disabledFromCtx = getCtxFlag(raw as Record<string, unknown>, [
    'isDisabled',
    'disabled',
    'formDisabled',
    'Disabled',
  ]);
  const AllDisabledFields = raw.AllDisabledFields;
  const AllHiddenFields = raw.AllHiddenFields;

  // Local state
  const [required, setRequired] = React.useState<boolean>(Boolean(isRequired));
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || Boolean(submitting) || isListed(AllDisabledFields, displayName)
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(isListed(AllHiddenFields, displayName));

  const [files, setFiles] = React.useState<File[]>([]);
  const [error, setError] = React.useState<string>('');

  // Existing SP attachments (Edit/View only, conditional)
  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | undefined>(undefined);
  const [loadingSP, setLoadingSP] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');

  const inputId = useId('file');
  const inputRef = React.useRef<HTMLInputElement>(null);

  // Single vs multi selection
  const isSingleSelection = !multiple || maxFiles === 1;

  /* ---------- effects ---------- */

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

    log('recompute flags:', {
      fromMode,
      fromCtx,
      fromSubmitting,
      fromDisabledList,
      fromHiddenList,
      isDisabled: fromMode || fromCtx || fromDisabledList || fromSubmitting,
      isHidden: fromHiddenList,
    });
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  // EDIT/VIEW: always attempt to fetch existing AttachmentFiles (single-item endpoint)
  React.useEffect((): void | (() => void) => {
    if (isNewMode) {
      logAtt('skip fetch: NEW mode');
      return;
    }

    const attachmentsFlag = readBool(FormData, 'attachments');
    logAtt('Edit/View: FormData.attachments =', attachmentsFlag);

    const { listTitle, itemId, shape } = getListTitleAndItemId(formCustomizerContext);
    logAtt('SPFx ctx read:', { shape, listTitle, itemId, rawCtx: formCustomizerContext });

    if (!listTitle || !itemId) {
      logAtt('skip fetch: missing listTitle or itemId');
      return;
    }

    // Single item endpoint => predictable shape { AttachmentFiles: [...] }
    const spUrl =
      `/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${encodeURIComponent(
        String(itemId)
      )})?$select=AttachmentFiles&$expand=AttachmentFiles`;

    let cancelled = false;

    (async (): Promise<void> => {
      setLoadingSP(true);
      setLoadError('');
      logAtt('fetch START:', spUrl);

      try {
        const respUnknown: unknown = await getFetchAPI({
          spUrl,
          method: 'GET',
          headers: { Accept: 'application/json;odata=nometadata' },
        });
        logAtt('fetch RESPONSE (raw):', respUnknown);

        const attsRaw = (respUnknown as { AttachmentFiles?: unknown } | null)?.AttachmentFiles;
        logAtt('fetch RESPONSE AttachmentFiles (raw):', attsRaw);

        const atts: SPAttachment[] = Array.isArray(attsRaw)
          ? attsRaw
              .map((x): SPAttachment | undefined => {
                if (x && typeof x === 'object') {
                  const o = x as Record<string, unknown>;
                  const FileName = typeof o.FileName === 'string' ? o.FileName : undefined;
                  const ServerRelativeUrl =
                    typeof o.ServerRelativeUrl === 'string' ? o.ServerRelativeUrl : undefined;
                  if (FileName && ServerRelativeUrl) return { FileName, ServerRelativeUrl };
                }
                return undefined;
              })
              .filter((x): x is SPAttachment => x !== undefined)
          : [];

        if (!cancelled) {
          setSpAttachments(atts);
          if (atts.length > 0) {
            logAtt(`fetch SUCCESS: ${atts.length} attachment(s) found`, atts);
          } else {
            logAtt('fetch SUCCESS: 0 attachments found');
          }
        }
      } catch (e: unknown) {
        const msg = e instanceof Error ? e.message : 'Failed to load attachments.';
        if (!cancelled) {
          setSpAttachments(undefined);
          setLoadError(msg);
          logAtt('fetch ERROR:', msg, e);
        }
      } finally {
        if (!cancelled) {
          setLoadingSP(false);
          logAtt('fetch FINISH');
        }
      }
    })().catch(err => {
      logAtt('fetch PROMISE catch (should not happen due to try/catch):', err);
    });

    return (): void => {
      cancelled = true;
      logAtt('effect cleanup: cancelled fetch');
    };
  }, [isNewMode, formCustomizerContext, FormData]);

  // 🔎 Anytime attachment state changes, log a clear status message
  React.useEffect((): void => {
    if (loadingSP) {
      logAtt('state change: loadingSP=true...');
      return;
    }
    if (loadError) {
      logAtt('state change: loadError ->', loadError);
      return;
    }
    if (Array.isArray(spAttachments)) {
      if (spAttachments.length > 0) {
        logAtt(
          `state change: ATTACHMENTS PRESENT (${spAttachments.length})`,
          spAttachments.map(a => ({ name: a.FileName, url: a.ServerRelativeUrl }))
        );
      } else {
        logAtt('state change: NO attachments (empty array)');
      }
    } else {
      logAtt('state change: attachments undefined (not loaded or fetch skipped/failed)');
    }
  }, [spAttachments, loadingSP, loadError]);

  /* ---------- validation & commit ---------- */

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

  const commitValue = React.useCallback(
    (list: File[]): void => {
      const payload = list.length === 0 ? undefined : isSingleSelection ? list[0] : list;
      log('commitValue -> GlobalFormData:', { id, payload });
      GlobalFormData(id, payload);
    },
    [GlobalFormData, id, isSingleSelection, log]
  );

  /* ---------- handlers ---------- */

  const openPicker = (): void => {
    log('openPicker click');
    if (!isDisabled) inputRef.current?.click();
    else log('openPicker prevented: isDisabled');
  };

  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e): void => {
    const picked = Array.from(e.currentTarget.files ?? []);
    log('onFilesPicked:', picked.map(f => ({ name: f.name, size: f.size, type: f.type })));

    let next: File[] = [];
    let msg = '';

    if (isSingleSelection) {
      next = picked.slice(0, 1);
      log('single selection -> taking first file only');
    } else {
      const already = files.length;
      const capacity = isDefined(maxFiles) ? Math.max(0, maxFiles - already) : picked.length;
      log('multi selection -> already:', already, 'capacity:', capacity);

      if (already === 0) {
        const toTake = isDefined(maxFiles) ? Math.min(picked.length, maxFiles) : picked.length;
        next = picked.slice(0, toTake);
        if (isDefined(maxFiles) && picked.length > maxFiles) {
          msg = TOO_MANY_MSG(maxFiles);
        }
      } else {
        const toAdd = picked.slice(0, capacity);
        next = files.concat(toAdd);
        if (isDefined(maxFiles) && picked.length > capacity) {
          msg = TOO_MANY_MSG(maxFiles);
        }
      }
    }

    if (!msg) msg = validateSelection(next);
    log('selection after apply:', next.map(f => f.name), 'validation msg:', msg);

    setFiles(next);
    setError(msg);
    GlobalErrorHandle(id, msg === '' ? undefined : msg);
    commitValue(next);

    // Allow selecting the same file(s) again
    if (inputRef.current) inputRef.current.value = '';
  };

  const removeAt = React.useCallback(
    (idx: number): void => {
      log('removeAt index:', idx);
      const next = files.filter((_, i) => i !== idx);
      const msg = validateSelection(next);

      setFiles(next);
      setError(msg);
      GlobalErrorHandle(id, msg === '' ? undefined : msg);
      commitValue(next);
    },
    [files, validateSelection, GlobalErrorHandle, id, commitValue, log]
  );

  const handleRemove = React.useCallback(
    (idx: number): React.MouseEventHandler<HTMLButtonElement> =>
      (): void => removeAt(idx),
    [removeAt]
  );

  const clearAll = (): void => {
    log('clearAll');
    const msg = required ? REQUIRED_MSG : '';
    setFiles([]);
    setError(msg);
    GlobalErrorHandle(id, msg || undefined);
    commitValue([]);
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ---------- render ---------- */

  if (isHidden) {
    log('render: hidden');
    return <div hidden className="fieldClass" />;
  }

  log('render: visible', { filesCount: files.length, error, isDisabled });

  return (
    <div className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {/* Existing attachments (Edit/View only) */}
        {!isNewMode && (
          <div style={{ marginBottom: 8 }}>
            {loadingSP && <Text size={200}>Loading attachments…</Text>}
            {!loadingSP && loadError && (
              <Text size={200} aria-live="polite">
                Error: {loadError}
              </Text>
            )}
            {!loadingSP && !loadError && Array.isArray(spAttachments) && spAttachments.length > 0 && (
              <>
                {logAtt('render: showing attachments list (count=%d)', spAttachments.length)}
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
                          <Link href={a.ServerRelativeUrl} target="_blank" rel="noreferrer">
                            {a.FileName}
                          </Link>
                        </div>
                        <Text size={200}>{a.ServerRelativeUrl}</Text>
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}
            {!loadingSP && !loadError && Array.isArray(spAttachments) && spAttachments.length === 0 && (
              <>
                {logAtt('render: no attachments to show (empty array)')}
                <Text size={200}>No existing attachments.</Text>
              </>
            )}
          </div>
        )}

        {/* Hidden native input (triggered by the button) */}
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

        {/* Actions */}
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

        {/* Locally selected files */}
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
