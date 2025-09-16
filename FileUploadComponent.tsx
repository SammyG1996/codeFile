/**
 * FileUploadComponent.tsx
 *
 * Summary
 * - Fluent UI v9 wrapper (<Field>) with a custom <input type="file"> trigger.
 * - Validations: required, accept (MIME/extensions), max file size, max files.
 * - Disabled = (FormMode===4) OR (context disabled flags) OR (AllDisabledFields) OR (submitting).
 * - Hidden  = (AllHiddenFields) — hides the wrapper <div>.
 *
 * Behavior
 * - No global writes on mount.
 * - Calls GlobalFormData only after user selects/clears files.
 * - GlobalErrorHandle is called only after first interaction (touched).
 * - Commits:
 *     · empty selection → null
 *     · otherwise → (single: File) | (multiple: File[])
 *
 * Notes
 * - This component does not upload to SharePoint by itself; it only surfaces selected File(s)
 *   via GlobalFormData so your parent flow can handle persistence.
 */

import * as React from 'react';
import { Field, Button, Text, useId } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, UploadRegular } from '@fluentui/react-icons';
import { DynamicFormContext } from './DynamicFormContext';

/* ---------- Props ---------- */

export interface FileUploadProps {
  id: string;
  displayName: string;

  /** If true, allow selecting multiple files */
  multiple?: boolean;

  /** Accept string (e.g., ".pdf,.docx,image/*" or "application/pdf") */
  accept?: string;

  /** Max file size in MB (per file). Omit for unlimited. */
  maxFileSizeMB?: number;

  /** Max number of files (when multiple=true). Omit for unlimited. */
  maxFiles?: number;

  /** Required selection? */
  isRequired?: boolean;

  /** Help text under the control */
  description?: string;

  /** Extra class for the trigger/input row */
  className?: string;

  /** While parent is submitting (disables UI) */
  submitting?: boolean;
}

/* ---------- Helpers ---------- */

const REQUIRED_MSG = 'Please select a file.';
const TOO_LARGE_MSG = (name: string, limitMB: number) =>
  `“${name}” exceeds the maximum size of ${limitMB} MB.`;
const TOO_MANY_MSG = (limit: number) =>
  `You can attach up to ${limit} file${limit === 1 ? '' : 's'}.`;

const formatBytes = (bytes: number): string => {
  if (!Number.isFinite(bytes)) return '';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let idx = 0;
  let val = bytes;
  while (val >= 1024 && idx < units.length - 1) {
    val /= 1024;
    idx++;
  }
  return `${val % 1 === 0 ? val.toFixed(0) : val.toFixed(2)} ${units[idx]}`;
};

// We treat "defined" as "not undefined" (avoid runtime null checks)
const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

/** Generic, safe access to unknown context shape without `any`. */
const hasKey = (obj: Record<string, unknown>, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(obj, key);
const getKey = <T,>(obj: Record<string, unknown>, key: string): T =>
  obj[key] as T;
/** Read a boolean-ish flag from one of several possible keys. */
const getCtxFlag = (obj: Record<string, unknown>, keys: string[]): boolean => {
  for (const k of keys) if (hasKey(obj, k)) return !!obj[k];
  return false;
};
/** Membership over unknown list-like: array, Set, comma string, or object map */
const toBool = (v: unknown): boolean => !!v;
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
    for (const [k, v] of Object.entries(bag as Record<string, unknown>)) {
      if (k.trim().toLowerCase() === needle && toBool(v)) return true;
    }
    return false;
  }
  return false;
};

/* ---------- Component ---------- */

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

  // Context (do not re-declare concrete shape locally)
  const formCtx = React.useContext(DynamicFormContext);
  const ctx = formCtx as unknown as Record<string, unknown>;

  const FormMode = hasKey(ctx, 'FormMode') ? getKey<number>(ctx, 'FormMode') : undefined;

  // These two must exist on the provider; assert-read
  const GlobalFormData = getKey<(id: string, value: unknown) => void>(ctx, 'GlobalFormData');
  const GlobalErrorHandle = getKey<(id: string, error: string | null) => void>(ctx, 'GlobalErrorHandle');

  const isDisplayForm = FormMode === 4;
  const disabledFromCtx = getCtxFlag(ctx, ['isDisabled', 'disabled', 'formDisabled', 'Disabled']);

  // Disabled/hidden lists (optional on context)
  const AllDisabledFields = hasKey(ctx, 'AllDisabledFields') ? ctx.AllDisabledFields : undefined;
  const AllHiddenFields = hasKey(ctx, 'AllHiddenFields') ? ctx.AllHiddenFields : undefined;

  // Flags
  const [required, setRequired] = React.useState<boolean>(!!isRequired);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || !!submitting || isListed(AllDisabledFields, displayName)
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(isListed(AllHiddenFields, displayName));

  // Value & validation state
  const [files, setFiles] = React.useState<File[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const inputId = useId('file');
  const inputRef = React.useRef<HTMLInputElement>(null);

  /* ---------- effects ---------- */

  React.useEffect(() => {
    setRequired(!!isRequired);
  }, [isRequired]);

  // Disabled/Hidden recompute
  React.useEffect(() => {
    const fromMode = isDisplayForm;
    const fromCtx = disabledFromCtx;
    const fromSubmitting = !!submitting;
    const fromDisabledList = isListed(AllDisabledFields, displayName);
    const fromHiddenList = isListed(AllHiddenFields, displayName);

    setIsDisabled(fromMode || fromCtx || fromDisabledList || fromSubmitting);
    setIsHidden(fromHiddenList);
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  // On mount: do NOT push anything
  React.useEffect(() => {
    // Intentionally no GlobalFormData/GlobalErrorHandle here.
  }, []);

  /* ---------- validation ---------- */

  const validateSelection = React.useCallback(
    (list: File[]): string => {
      if (required && list.length === 0) return REQUIRED_MSG;
      if (multiple && isDefined(maxFiles) && list.length > maxFiles) return TOO_MANY_MSG(maxFiles);

      if (isDefined(maxFileSizeMB)) {
        const limit = maxFileSizeMB * 1024 * 1024;
        for (const f of list) {
          if (f.size > limit) return TOO_LARGE_MSG(f.name, maxFileSizeMB);
        }
      }
      // NOTE: <input accept> already filters chooser UI, but we keep the string for the element.
      return '';
    },
    [required, multiple, maxFiles, maxFileSizeMB]
  );

  const commitValue = React.useCallback(
    (list: File[]) => {
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(id, list.length === 0 ? null : (multiple ? list : list[0]));
    },
    [GlobalFormData, id, multiple]
  );

  const pushErrorIfTouched = React.useCallback(
    (msg: string) => {
      setError(msg);
      if (touched) {
        // eslint-disable-next-line @rushstack/no-new-null
        GlobalErrorHandle(id, msg === '' ? null : msg);
      }
    },
    [GlobalErrorHandle, id, touched]
  );

  /* ---------- handlers ---------- */

  const openPicker = (): void => {
    if (isDisabled) return;
    inputRef.current?.click();
  };

  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = (e): void => {
    setTouched(true);

    const list = Array.from(e.currentTarget.files ?? []);
    const msg = validateSelection(list);

    setFiles(list);
    setError(msg);

    // Update global error immediately after first interaction
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg === '' ? null : msg);

    // Commit value regardless; parent can decide what to do with an empty/null or with error present
    commitValue(list);
  };

  const removeAt = (idx: number): void => {
    setTouched(true);
    const next = files.filter((_, i) => i !== idx);
    const msg = validateSelection(next);
    setFiles(next);
    setError(msg);
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, msg === '' ? null : msg);
    commitValue(next);
  };

  const clearAll = (): void => {
    setTouched(true);
    setFiles([]);
    setError(required ? REQUIRED_MSG : '');
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(id, (required ? REQUIRED_MSG : '') || null);
    commitValue([]);
    // Clear input so picking the same file again will fire onChange
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ---------- render ---------- */

  return (
    <div hidden={isHidden} className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
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
            icon={<UploadRegular />}
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

          {/* Lightweight “requirements” hint */}
          {(accept || maxFileSizeMB || (multiple && maxFiles)) && (
            <Text size={200} wrap>
              {accept && <span>Allowed: <code>{accept}</code>. </span>}
              {isDefined(maxFileSizeMB) && <span>Max size: {maxFileSizeMB} MB/file. </span>}
              {multiple && isDefined(maxFiles) && <span>Max files: {maxFiles}.</span>}
            </Text>
          )}
        </div>

        {/* Selected files list */}
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
