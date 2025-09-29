/**
 * FileUploadComponent.tsx (clean build: no console logs, no `any`, no floating promises, no `void` op, dot-notation)
 *
 * Example:
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

/** Interpret DynamicFormContext.FormData.Attachments (or similar) as a boolean hint. */
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
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  // EDIT/VIEW: fetch existing AttachmentFiles only if FormData indicates there ARE attachments
  React.useEffect((): void | (() => void) => {
    if (isNewMode) return;

    const attachmentsHint = readAttachmentsHint(FormData);
    if (attachmentsHint === false) {
      setSpAttachments([]);
      setLoadingSP(false);
      setLoadError('');
      return;
    }

    // Pull identifiers from SPFx context (safely)
    const listTitle: string | undefined = (context as { list?: { title?: string } } | undefined)?.list?.title;
    const listGuid: string | undefined = (context as { list?: { id?: string } } | undefined)?.list?.id;
    const itemId: number | undefined = (context as { item?: { ID?: number } } | undefined)?.item?.ID;
    const baseUrl: string | undefined =
      (context as { pageContext?: { web?: { absoluteUrl?: string } } } | undefined)?.pageContext?.web?.absoluteUrl ??
      (typeof window !== 'undefined' ? window.location.origin : undefined);

    if (!baseUrl || !itemId || (!listGuid && !listTitle)) {
      return;
    }

    const encTitle = listTitle ? encodeURIComponent(listTitle) : '';
    const idStr = encodeURIComponent(String(itemId));

    // Absolute URL candidates (with and without "/_api")
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
              // dot-notation-friendly cast
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
          // continue to next candidate
        }
      }

      if (!success && !cancelled) {
        const msg = lastErr instanceof Error ? lastErr.message : 'Failed to load attachments.';
        setSpAttachments(undefined);
        setLoadError(msg);
        setLoadingSP(false);
      }
    })().catch(() => {
      // Suppress unhandled rejection to satisfy eslint rules without using `void` operator.
    });

    return (): void => {
      cancelled = true;
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
      const payload: unknown = list.length === 0 ? undefined : isSingleSelection ? list[0] : list;
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
    raw.GlobalErrorHandle(id, msg === '' ? undefined : msg);
    commitValue(next);

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