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
import { Field, Button, Text, Link } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';
import { DynamicFormContext } from './DynamicFormContext';

// 👇 YOUR app's React context that exposes SPFx-like info (list, item, etc.)
//    Update this import path to wherever your provider lives.
import { FormCustomizerContext } from '../contexts/FormCustomizerContext';

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

/**
 * Robustly infer whether FormData indicates the item has attachments.
 * Checks common shapes and returns BOTH the boolean and what key triggered it.
 */
const readAttachmentsFlagLoose = (
  obj: unknown
): { value: boolean; source: 'attachments' | 'Attachments' | 'AttachmentFiles' | 'attachmentFiles' | 'count' | 'none' } => {
  if (!obj || typeof obj !== 'object') return { value: false, source: 'none' };
  const o = obj as Record<string, unknown>;

  if (typeof o.attachments === 'boolean') return { value: o.attachments, source: 'attachments' };
  if (typeof o.Attachments === 'boolean') return { value: o.Attachments, source: 'Attachments' };

  if (Array.isArray(o.AttachmentFiles)) return { value: o.AttachmentFiles.length > 0, source: 'AttachmentFiles' };
  if (Array.isArray(o.attachmentFiles)) return { value: o.attachmentFiles.length > 0, source: 'attachmentFiles' };

  if (typeof o.attachments === 'number') return { value: (o.attachments as number) > 0, source: 'count' };
  if (typeof o.Attachments === 'number') return { value: (o.Attachments as number) > 0, source: 'count' };

  return { value: false, source: 'none' };
};

/**
 * Tries VERY HARD to find listTitle and itemId across BOTH contexts and FormData.
 * Priority:
 *   1) Your dedicated React FormCustomizerContext (wrapped & direct shapes)
 *   2) DynamicFormContext (some apps mirror values here)
 *   3) FormData (ID often lives here)
 */
const getListTitleAndItemIdLoose = (
  fcCtxValue: unknown,            // value from your FormCustomizerContext
  dynamicCtxValue: unknown,       // the whole DynamicFormContext object
  formData: unknown               // FormData from DynamicFormContext
): {
  listTitle?: string;
  itemId?: number;
  source: string;
} => {
  const pickNum = (v: unknown): number | undefined => {
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    if (typeof v === 'string' && /^\d+$/.test(v)) return Number(v);
    return undefined;
  };
  const pickStr = (v: unknown): string | undefined =>
    typeof v === 'string' && v.trim() ? v : undefined;

  const tryPaths = (obj: unknown, paths: string[]): unknown => {
    if (!obj || typeof obj !== 'object') return undefined;
    for (const p of paths) {
      const parts = p.split('.');
      let cur: any = obj;
      let ok = true;
      for (const part of parts) {
        if (cur && typeof cur === 'object' && part in cur) cur = cur[part];
        else { ok = false; break; }
      }
      if (ok) return cur;
    }
    return undefined;
  };

  // 1) Your dedicated FormCustomizerContext (wrapped + direct)
  const titleFromFC = pickStr(
    tryPaths(fcCtxValue, [
      'context.list.title', // wrapped SPFx-like
      'list.title',         // direct
      'context.listTitle',
      'listTitle',
    ])
  );
  const idFromFC = pickNum(
    tryPaths(fcCtxValue, [
      'context.item.ID',
      'item.ID',
      'context.itemId',
      'itemId',
      'context.item.Id',
      'item.Id',
    ])
  );
  if (titleFromFC && idFromFC !== undefined) {
    return { listTitle: titleFromFC, itemId: idFromFC, source: 'FormCustomizerContext' };
  }

  // 2) DynamicFormContext mirror (some apps copy values here)
  const titleFromDyn = pickStr(
    tryPaths(dynamicCtxValue, ['list.title', 'listTitle', 'ListTitle', 'Title'])
  );
  const idFromDyn = pickNum(
    tryPaths(dynamicCtxValue, ['item.ID', 'itemId', 'ItemID', 'ID', 'Id'])
  );
  if (titleFromDyn && idFromDyn !== undefined) {
    return { listTitle: titleFromDyn, itemId: idFromDyn, source: 'DynamicFormContext' };
  }

  // 3) FormData fallback (ID is often present)
  const idFromFD = pickNum(tryPaths(formData, ['ID', 'Id']));
  const titleFromFD = pickStr(tryPaths(formData, ['ListTitle', 'listTitle']));
  if (titleFromFD && idFromFD !== undefined) {
    return { listTitle: titleFromFD, itemId: idFromFD, source: 'FormData' };
  }
  if (idFromFD !== undefined && titleFromFC) {
    return { listTitle: titleFromFC, itemId: idFromFD, source: 'mixed: FC.title + FormData.ID' };
  }

  return { source: 'not-found' };
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
    console.log(
      '%c[%cFileUpload%c] %cATTACHMENTS',
      'color:#888',
      'color:#0b6',
      'color:#888',
      'color:#06c;font-weight:bold',
      ...args
    );
  };

  log('mount: props =', { id, displayName, multiple, accept, maxFileSizeMB, maxFiles, isRequired, submitting });

  // Contexts
  const raw = React.useContext(DynamicFormContext) as unknown as FormCtxShape;
  // If your FormCustomizerContext is typed, you can add the generic instead of `unknown`
  const fcValue = React.useContext(FormCustomizerContext as unknown as React.Context<unknown>);

  log('context snapshot (DynamicFormContext):', raw);
  log('context snapshot (FormCustomizerContext):', fcValue);

  const FormData = raw.FormData;
  const FormMode = raw.FormMode;
  const GlobalFormData = raw.GlobalFormData as (id: string, value: unknown) => void;
  const GlobalErrorHandle = raw.GlobalErrorHandle as (id: string, error: string | undefined) => void;

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

  // EDIT/VIEW: fetch existing AttachmentFiles using COLLECTION endpoint with $filter
  React.useEffect((): void | (() => void) => {
    // 1) Mode must be Edit/View (not New)
    if (isNewMode) {
      logAtt('fetch CONDITIONS:', { isNewMode, listTitle: undefined, itemId: undefined, source: '—', canFetch: false });
      logAtt('fetch SKIPPED (conditions not met): isNewMode === true (New form)');
      return;
    }

    // 2) FormData hint (just for visibility)
    const { value: fdHasAttachments, source: fdSource } = readAttachmentsFlagLoose(FormData);
    logAtt('FormData attachments flag (loose):', { value: fdHasAttachments, source: fdSource });
    logAtt('FormData keys:', FormData ? Object.keys(FormData) : 'no FormData');

    // 3) Need both listTitle and itemId (from EITHER context, or FormData fallback)
    const { listTitle, itemId, source } = getListTitleAndItemIdLoose(
      fcValue,
      raw,
      FormData
    );

    const canFetch = Boolean(!isNewMode && listTitle && itemId);
    logAtt('fetch CONDITIONS:', { isNewMode, listTitle, itemId, source, canFetch });

    if (!canFetch) {
      logAtt('fetch SKIPPED (conditions not met):', {
        reason_isNewMode: isNewMode,
        reason_missingListTitle: !listTitle,
        reason_missingItemId: !itemId,
      });
      return;
    }

    logAtt('fetch TRIGGERED: all conditions satisfied');

    // ✅ Matches instructions (collection + $filter + $select/$expand)
    const spUrl =
      `/_api/web/lists/getbytitle('${encodeURIComponent(listTitle as string)}')/items` +
      `?$filter=Id eq ${encodeURIComponent(String(itemId))}` +
      `&$select=AttachmentFiles&$expand=AttachmentFiles`;

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

        // Collection shape: { value: [ { AttachmentFiles: [...] } ] }
        const rows = ((respUnknown as { value?: unknown[] } | null)?.value ?? []) as unknown[];
        const firstRow = Array.isArray(rows) ? (rows[0] as { AttachmentFiles?: unknown } | undefined) : undefined;
        const attsRaw = firstRow?.AttachmentFiles;
        logAtt('fetch RESPONSE firstRow.AttachmentFiles (raw):', attsRaw);

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
  }, [isNewMode, raw, fcValue, FormData]);

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
          id={id}            /* use id from props */
          name={displayName} /* set name to displayName */
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