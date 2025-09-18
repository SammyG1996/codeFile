/**
 * Example usage:
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
 * />
 *
 * ——————————————————————————————————————————————————————————————————————
 *
 * FileUploadComponent.tsx
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
}

/** Minimal, type-safe view of our form context. All fields optional. */
type FormCtxShape = {
  FormData?: Record<string, unknown>;
  FormMode?: number;
  GlobalFormData?: (id: string, value: unknown) => void;
  GlobalErrorHandle?: (id: string, error: string | null) => void;

  // The SPFx form context instance (may be the context itself or wrapped under `.context`)
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
  if (bag == null) return false;

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
const getListTitleAndItemId = (ctx: unknown): { listTitle?: string; itemId?: number } => {
  const root =
    ctx && typeof ctx === 'object' && hasKey(ctx as Record<string, unknown>, 'context')
      ? (ctx as { context: unknown }).context
      : ctx;

  if (!root || typeof root !== 'object') return {};

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

  return { listTitle, itemId };
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
  } = props;

  // Context (typed to our minimal shape)
  const raw = React.useContext(DynamicFormContext) as unknown as FormCtxShape;

  const FormData = raw.FormData;
  const FormMode = raw.FormMode;
  const GlobalFormData = raw.GlobalFormData as (id: string, value: unknown) => void;
  const GlobalErrorHandle = raw.GlobalErrorHandle as (id: string, error: string | null) => void;

  // Treat this as unknown, we’ll safely read the fields we need
  const formCustomizerContext: unknown = raw.FormCustomizerContext as unknown as SPFxFormCustomizerContext | unknown;

  const isDisplayForm = FormMode === 4; // VIEW
  const isNewMode = FormMode === 8; // NEW

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
  const [spAttachments, setSpAttachments] = React.useState<SPAttachment[] | null>(null);
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
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  // EDIT/VIEW: fetch existing AttachmentFiles ONLY if FormData.attachments === true
  React.useEffect((): void | (() => void) => {
    if (isNewMode) return;

    const hasAttachmentsFlag = readBool(FormData, 'attachments');
    if (!hasAttachmentsFlag) {
      setSpAttachments(null);
      setLoadError('');
      return;
    }

    const { listTitle, itemId } = getListTitleAndItemId(formCustomizerContext);
    if (!listTitle || !itemId) return;

    const spUrl =
      `/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items` +
      `?$filter=Id eq ${encodeURIComponent(String(itemId))}` +
      `&$select=AttachmentFiles&$expand=AttachmentFiles`;

    let cancelled = false;

    // Use `void` to satisfy no-floating-promises and handle errors internally.
    void (async (): Promise<void> => {
      setLoadingSP(true);
      setLoadError('');
      try {
        const respUnknown: unknown = await getFetchAPI({
          spUrl,
          method: 'GET',
          headers: { Accept: 'application/json;odata=nometadata' },
        });

        // Narrow the response safely
        const rows: unknown = (respUnknown as { value?: unknown[] } | null)?.value ?? [];
        const firstRow = Array.isArray(rows) ? rows[0] : undefined;
        const attsRaw: unknown =
          firstRow && typeof firstRow === 'object' ? (firstRow as Record<string, unknown>)['AttachmentFiles'] : null;

        const atts: SPAttachment[] = Array.isArray(attsRaw)
          ? attsRaw
              .map((x): SPAttachment | null => {
                if (x && typeof x === 'object') {
                  const o = x as Record<string, unknown>;
                  const FileName = typeof o.FileName === 'string' ? o.FileName : '';
                  const ServerRelativeUrl = typeof o.ServerRelativeUrl === 'string' ? o.ServerRelativeUrl : '';
                  if (FileName && ServerRelativeUrl) return { FileName, ServerRelativeUrl };
                }
                return null;
              })
              .filter((x): x is SPAttachment => x !== null)
          : [];

        if (!cancelled) setSpAttachments(atts);
      } catch (e: unknown) {
        const msg = e instanceof Error ? e.message : 'Failed to load attachments.';
        if (!cancelled) {
          setSpAttachments(null);
          setLoadError(msg);
        }
      } finally {
        if (!cancelled) setLoadingSP(false);
      }
    })();

    // ✅ Explicitly type the cleanup return
    return (): void => {
      cancelled = true;
    };
  }, [isNewMode, formCustomizerContext, FormData]);

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
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(id, list.length === 0 ? null : isSingleSelection ? list[0] : list);
    },
    [GlobalFormData, id, isSingleSelection]
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

    setFiles(next);
    setError(msg);
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg === '' ? null : msg);
    commitValue(next);

    // Allow selecting the same file(s) again
    if (inputRef.current) inputRef.current.value = '';
  };

  const removeAt = React.useCallback(
    (idx: number): void => {
      const next = files.filter((_, i) => i !== idx);
      const msg = validateSelection(next);

      setFiles(next);
      setError(msg);
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalErrorHandle(id, msg === '' ? null : msg);
      commitValue(next);
    },
    [files, validateSelection, GlobalErrorHandle, id, commitValue]
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
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg || null);
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
        {/* Existing attachments (Edit/View only, and only if item has attachments) */}
        {!isNewMode && (
          <div style={{ marginBottom: 8 }}>
            {loadingSP && <Text size={200}>Loading attachments…</Text>}
            {!loadingSP && loadError && (
              <Text size={200} aria-live="polite">
                Error: {loadError}
              </Text>
            )}
            {!loadingSP && !loadError && spAttachments && spAttachments.length > 0 && (
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
            {!loadingSP && !loadError && spAttachments && spAttachments.length === 0 && (
              <Text size={200}>No existing attachments.</Text>
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