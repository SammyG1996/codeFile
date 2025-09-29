/**
 * FileUploadComponent.tsx (diagnostics + absolute SP URLs)
 *
 * Example usage:
 * <FileUploadComponent
 *   id="Attachments"
 *   displayName="Attachments"
 *   multiple
 *   accept=".pdf,.doc,.docx,image/*"
 *   maxFiles={5}
 *   maxFileSizeMB={15}
 *   isRequired={false}
 *   description="Add any supporting files."
 *   submitting={isSubmitting}
 *   context={props.context} // SPFx FormCustomizerContext
 * />
 */

import * as React from 'react';
import { Field, Button, Text, Link } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';
import { DynamicFormContext } from './DynamicFormContext';
import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
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

  /** SPFx Form Customizer context – used to build the REST URL */
  context?: FormCustomizerContext;
}

type FormCtxShape = {
  FormData?: Record<string, unknown>;
  FormMode?: number;
  GlobalFormData: (id: string, value: unknown) => void;
  GlobalErrorHandle: (id: string, error: string | undefined) => void;

  // optional flags/lists
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
  while (n >= 1024 && i < units.length - 1) {
    n /= 1024;
    i++;
  }
  return `${Number.isInteger(n) ? n.toFixed(0) : n.toFixed(2)} ${units[i]}`;
};

/** Interpret DynamicFormContext.FormData.Attachments as a boolean hint. */
const readAttachmentsHint = (fd: Record<string, unknown> | undefined): boolean | undefined => {
  if (!fd) return undefined;
  const v =
    (fd as any).Attachments ??
    (fd as any).attachments ??
    (fd as any).AttachmentCount ??
    (fd as any).attachmentCount;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v > 0;
  return undefined;
};

/* Pretty, scoped logger */
const tag = (label: string): string => `%c[%cFileUpload%c] %c${label}`;
const base = 'color:#888';
const hi = 'color:#0b6';
const lab = 'color:#06c;font-weight:bold';

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
    context, // SPFx FormCustomizerContext
  } = props;

  // Dynamic form context (supplies modes + commit hooks)
  const raw = React.useContext(DynamicFormContext) as unknown as FormCtxShape;

  const FormData = raw.FormData;
  const FormMode = raw.FormMode ?? 0;

  // Modes
  const isDisplayForm = FormMode === 4; // VIEW
  const isNewMode = FormMode === 8; // NEW

  // Disabled/hidden from context/passed flags
  const disabledFromCtx = getCtxFlag(raw as unknown as Record<string, unknown>, [
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

  // Existing SP attachments (Edit/View only when FormData says there are attachments)
  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | undefined>(undefined);
  const [loadingSP, setLoadingSP] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');

  const inputRef = React.useRef<HTMLInputElement>(null);

  // Single vs multi selection
  const isSingleSelection = !multiple || maxFiles === 1;

  /* ---------- initial diagnostics ---------- */

  React.useEffect(() => {
    // eslint-disable-next-line no-console
    console.log(tag('MOUNT'), base, hi, base, lab, {
      props: { id, displayName, multiple, accept, maxFileSizeMB, maxFiles, isRequired, submitting },
      formMode: FormMode,
      formDataKeys: FormData ? Object.keys(FormData) : '(no FormData)',
      contextSummary: {
        hasContext: Boolean(context),
        listTitle: context?.list?.title,
        listId: (context as any)?.list?.id,
        itemId: context?.item?.ID,
        webUrl: (context as any)?.pageContext?.web?.absoluteUrl,
      },
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

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

    // eslint-disable-next-line no-console
    console.log(tag('FLAGS recompute'), base, hi, base, lab, {
      fromMode,
      fromCtx,
      fromSubmitting,
      fromDisabledList,
      fromHiddenList,
      computed: {
        isDisabled: fromMode || fromCtx || fromDisabledList || fromSubmitting,
        isHidden: fromHiddenList,
      },
    });
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  // EDIT/VIEW: fetch existing AttachmentFiles only if FormData indicates there ARE attachments
  React.useEffect((): void | (() => void) => {
    if (isNewMode) {
      console.log(tag('FETCH skip'), base, hi, base, lab, 'New form (FormMode === 8)');
      return;
    }

    const attachmentsHint = readAttachmentsHint(FormData);
    console.log(tag('FormData.Attachments'), base, hi, base, lab, attachmentsHint, { FormData });
    if (attachmentsHint === false) {
      console.log(tag('FETCH skip'), base, hi, base, lab, 'FormData indicates no attachments');
      setSpAttachments([]);
      setLoadingSP(false);
      setLoadError('');
      return;
    }

    const listTitle = context?.list?.title as string | undefined;
    const listGuid = (context as any)?.list?.id as string | undefined; // if available
    const itemId = context?.item?.ID as number | undefined;
    const baseUrl =
      ((context as any)?.pageContext?.web?.absoluteUrl as string | undefined) ??
      (typeof window !== 'undefined' ? window.location.origin : '');

    console.log(tag('FETCH conditions'), base, hi, base, lab, {
      baseUrl,
      listTitle,
      listGuid,
      itemId,
      canFetch: Boolean(baseUrl && itemId && (listGuid || listTitle)),
    });

    if (!baseUrl || !itemId || (!listGuid && !listTitle)) {
      console.log(tag('FETCH skip'), base, hi, base, lab, 'Missing baseUrl, itemId, or list identifier');
      return;
    }

    const encTitle = listTitle ? encodeURIComponent(listTitle) : '';
    const idStr = encodeURIComponent(String(itemId));

    // Build absolute URL candidates (some helpers want /_api, some add it; support both).
    const urls: (string | undefined)[] = [];

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

    const candidates = urls.filter(Boolean) as string[];

    let cancelled = false;

    (async (): Promise<void> => {
      setLoadingSP(true);
      setLoadError('');

      let lastErr: unknown = null;

      for (const spUrl of candidates) {
        if (cancelled) return;

        console.log(tag('FETCH try ->'), base, hi, base, lab, spUrl);
        try {
          const respUnknown: unknown = await getFetchAPI({
            spUrl,
            method: 'GET',
            headers: { Accept: 'application/json;odata=nometadata' },
          });

          console.log(tag('FETCH response (raw)'), base, hi, base, lab, { spUrl, respUnknown });

          // Shape can be one item or collection
          let attsRaw: unknown;
          if (respUnknown && typeof respUnknown === 'object') {
            const r = respUnknown as Record<string, unknown>;
            if (Array.isArray(r.value)) {
              const first = r.value[0] as Record<string, unknown> | undefined;
              attsRaw = first?.AttachmentFiles;
            } else if (r.AttachmentFiles !== undefined) {
              attsRaw = r.AttachmentFiles;
            }
          }

          const atts: SPAttachment[] = Array.isArray(attsRaw)
            ? (attsRaw as unknown[]).map((x) => {
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
          console.log(tag('FETCH success'), base, hi, base, lab, `${atts.length} attachment(s)`, atts);
          setLoadingSP(false);
          return; // ✅ first success wins
        } catch (e) {
          lastErr = e;
          console.log(tag('FETCH failed'), base, hi, base, lab, { spUrl, error: e });
          // try next…
        }
      }

      // all failed
      if (!cancelled) {
        const msg = lastErr instanceof Error ? lastErr.message : 'Failed to load attachments.';
        setSpAttachments(undefined);
        setLoadError(msg);
        setLoadingSP(false);
        console.log(tag('FETCH error – all variants failed'), base, hi, base, lab, msg);
      }
    })();

    return (): void => {
      cancelled = true;
      console.log(tag('EFFECT cleanup'), base, hi, base, lab, 'cancelled=true');
    };
  }, [isNewMode, FormData, context]);

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
      console.log(tag('COMMIT GlobalFormData'), base, hi, base, lab, { id, payload });
      raw.GlobalFormData(id, payload);
    },
    [raw, id, isSingleSelection]
  );

  /* ---------- handlers ---------- */

  const openPicker = (): void => {
    if (!isDisabled) inputRef.current?.click();
  };

  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e): void => {
    const picked = Array.from(e.currentTarget.files ?? []);
    console.log(tag('PICKED input files'), base, hi, base, lab, picked.map(f => ({ name: f.name, size: f.size, type: f.type })));

    let next: File[] = [];
    let msg = '';

    if (isSingleSelection) {
      next = picked.slice(0, 1); // single: take the first file only
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
    raw.GlobalErrorHandle(id, msg === '' ? undefined : msg);
    commitValue(next);

    console.log(tag('STATE files set'), base, hi, base, lab, {
      count: next.length,
      names: next.map(f => f.name),
      error: msg || '(none)',
    });

    if (inputRef.current) inputRef.current.value = '';
  };

  const removeAt = React.useCallback(
    (idx: number): void => {
      const next = files.filter((_, i) => i !== idx);
      const msg = validateSelection(next);

      setFiles(next);
      setError(msg);
      raw.GlobalErrorHandle(id, msg === '' ? undefined : msg);
      commitValue(next);

      console.log(tag('REMOVE file'), base, hi, base, lab, { removedIndex: idx, remaining: next.map(f => f.name), error: msg || '(none)' });
    },
    [files, validateSelection, raw, id, commitValue]
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
    raw.GlobalErrorHandle(id, msg || undefined);
    commitValue([]);

    console.log(tag('CLEAR all files'), base, hi, base, lab, { error: msg || '(none)' });

    if (inputRef.current) inputRef.current.value = '';
  };

  /* ---------- render ---------- */

  if (isHidden) return <div hidden className="fieldClass" />;

  return (
    <div className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {/* Existing attachments (Edit/View only and only if hint allows) */}
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

        {/* Hidden native input (triggered by the button) */}
        <input
          id={id}            /* use prop id */
          name={displayName} /* name shows as displayName */
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
          <Button appearance="primary" icon={<AttachRegular />} onClick={(): void => openPicker()} disabled={isDisabled}>
            {files.length === 0
              ? isSingleSelection
                ? 'Choose file'
                : 'Choose files'
              : isSingleSelection
              ? 'Choose different file'
              : 'Add more files'}
          </Button>

          {files.length > 0 && (
            <Button appearance="secondary" onClick={(): void => clearAll()} icon={<DismissRegular />} disabled={isDisabled}>
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