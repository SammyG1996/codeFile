/**
 * FileUploadComponent.tsx
 */

import * as React from 'react';
import { Field, Button, Text, Link, Spinner } from '@fluentui/react-components';
import { DismissRegular, DocumentRegular, AttachRegular } from '@fluentui/react-icons';

import { DynamicFormContext } from './DynamicFormContext';
import type { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { getFetchAPI } from '../Utils/getFetchApi';
import { evaluateFieldRules } from '../Utils/formRulesEngine';

/* ------------------------------ Types ------------------------------ */

export interface FileUploadProps {
  id: string;
  displayName: string;
  multiple?: boolean; // default true
  isRequired?: boolean;
  description?: string;
  className?: string;
  submitting?: boolean;
  context?: FormCustomizerContext; // SPFx context for reading/deleting attachments
}

type SPAttachment = { FileName: string; ServerRelativeUrl: string };

// Narrow shape we use from DynamicFormContext
type FormCtxShape = {
  FormData: Record<string, unknown>;
  FormMode: number;
  GlobalFormData: (id: string, value: unknown) => void;
  GlobalErrorHandle: (id: string, error: string | undefined) => void;
  isDisabled?: boolean;
  disabled?: boolean;
  formDisabled?: boolean;
  Disabled?: boolean;
  AllDisabledFields?: unknown;
  AllHiddenFields?: unknown;
  curUserInfo: {};
  formRules: any;
};

/* ------------------------------ Constants ------------------------------ */

const REQUIRED_MSG = 'Please select at least one file.';
const TOTAL_LIMIT_MSG = 'Selected files exceed the 250 MB total size limit.';
const TOTAL_LIMIT_BYTES = 250 * 1024 * 1024; // 250 MB
const MAX_NAME_LEN = 150;

/* ------------------------------ Utilities ------------------------------ */

const getCtxFlag = (o: Record<string, unknown>, keys: string[]): boolean =>
  keys.some((k) => Object.prototype.hasOwnProperty.call(o, k) && Boolean(o[k]));

const isListed = (bag: unknown, name: string): boolean => {
  const needle = name.trim().toLowerCase();
  if (bag === null || bag === undefined) return false;

  if (Array.isArray(bag)) return bag.some((v) => String(v).trim().toLowerCase() === needle);

  if (typeof (bag as { has?: unknown }).has === 'function') {
    for (const v of bag as Set<unknown>) {
      if (String(v).trim().toLowerCase() === needle) return true;
    }
    return false;
  }

  if (typeof bag === 'string') {
    return bag
      .split(',')
      .map((s) => s.trim().toLowerCase())
      .includes(needle);
  }

  if (typeof bag === 'object') {
    for (const [k, v] of Object.entries(bag as Record<string, unknown>)) {
      if (k.trim().toLowerCase() === needle && Boolean(v)) return true;
    }
  }

  return false;
};

// Allow: letters, digits, space, underscore; dots as extension separators
const regexAllowedName = /^[A-Za-z0-9_ ]+(\.[A-Za-z0-9_ ]+)*$/;
const validFileName = (name: string): boolean => regexAllowedName.test(name);

// Pretty bytes
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

// Explicitly convert a File into a Blob (preserving MIME type)
function fileToBlob(file: File): Blob {
  return file.slice(0, file.size, file.type || 'application/octet-stream');
}

const normalizeName = (name: string): string => name.trim().toLowerCase();

/**
 * Filter a batch of newly picked files:
 * - Skips disallowed characters
 * - Skips names > MAX_NAME_LEN
 * - Skips duplicates against:
 *   a) already-selected files (existingNames)
 *   b) earlier files in the same picked batch
 * Returns only the valid/unique files + a single local warning string listing offending filenames.
 */
function filterNewFiles(
  picked: File[],
  existingNames: Set<string>
): { validNew: File[]; warning: string } {
  const seen = new Set(existingNames); // start with existing names
  const validNew: File[] = [];
  const badNames: string[] = [];
  const tooLong: string[] = [];
  const duplicates: string[] = [];

  for (const f of picked) {
    const name = f.name ?? '';
    const norm = normalizeName(name);

    if (name.length > MAX_NAME_LEN) {
      tooLong.push(name);
      continue;
    }

    if (!validFileName(name)) {
      badNames.push(name);
      continue;
    }

    if (seen.has(norm)) {
      duplicates.push(name);
      continue;
    }

    // accept and record
    validNew.push(f);
    seen.add(norm);
  }

  const parts: string[] = [];

  if (badNames.length > 0) {
    parts.push(
      `Skipped invalid filename${badNames.length > 1 ? 's' : ''}: ${badNames
        .map((n) => `"${n}"`)
        .join(', ')}.`
    );
  }

  if (tooLong.length > 0) {
    parts.push(
      `Skipped over-length filename${tooLong.length > 1 ? 's' : ''} (>${MAX_NAME_LEN} characters): ${tooLong
        .map((n) => `"${n}"`)
        .join(', ')}.`
    );
  }

  if (duplicates.length > 0) {
    parts.push(
      `Skipped duplicate filename${duplicates.length > 1 ? 's' : ''}: ${duplicates
        .map((n) => `"${n}"`)
        .join(', ')}.`
    );
  }

  return { validNew, warning: parts.join(' ') };
}

/**
 * Ensure filenames are unique by base name (case-insensitive).
 * Example:
 * - test.docx
 * - test.png
 * becomes:
 * - test.docx
 * - test_1.png
 *
 * Existing numeric suffixes are respected so a new duplicate does not
 * accidentally create another collision.
 */
function ensureUniqueFilenames(
  pickedFiles: File[],
  existingFullNames: Set<string>
): File[] {
  const usedBaseNames = new Set<string>();

  const getNameParts = (fileName: string): { base: string; extension: string } => {
    const lastDot = fileName.lastIndexOf('.');
    if (lastDot <= 0) {
      return { base: fileName, extension: '' };
    }

    return {
      base: fileName.slice(0, lastDot),
      extension: fileName.slice(lastDot),
    };
  };

  const splitBaseAndSuffix = (baseName: string): { root: string; suffixNumber: number | undefined } => {
    const match = baseName.match(/^(.*?)(?:_(\d+))?$/);
    if (!match) {
      return { root: baseName, suffixNumber: undefined };
    }

    return {
      root: match[1],
      suffixNumber: match[2] !== undefined ? Number(match[2]) : undefined,
    };
  };

  // Seed used base names from anything already selected / already attached.
  for (const fullName of existingFullNames) {
    const { base } = getNameParts(fullName);
    usedBaseNames.add(base.toLowerCase());
  }

  const renamedFiles: File[] = [];

  for (const file of pickedFiles) {
    const { base, extension } = getNameParts(file.name);
    const { root } = splitBaseAndSuffix(base);

    let nextBase = base;
    let counter = 1;

    while (usedBaseNames.has(nextBase.toLowerCase())) {
      nextBase = `${root}_${counter}`;
      counter += 1;
    }

    usedBaseNames.add(nextBase.toLowerCase());

    const nextName = `${nextBase}${extension}`;

    if (nextName === file.name) {
      renamedFiles.push(file);
    } else {
      renamedFiles.push(new File([file], nextName, { type: file.type }));
    }
  }

  return renamedFiles;
}

/* ------------------------------ Component ------------------------------ */

export default function FileUploadComponent(props: FileUploadProps): JSX.Element {
  const {
    id,
    displayName,
    multiple = true,
    isRequired,
    description = '',
    className,
    submitting,
    context,
  } = props;

  const ctx = DynamicFormContext() as unknown as FormCtxShape;

  const FormData = ctx.FormData;
  const FormMode = ctx.FormMode ?? 0;
  const isNewMode = FormMode === 8; // 8 = New
  const isDisplayForm = FormMode === 4; // 4 = Display

  const disabledFromCtx = getCtxFlag(ctx as unknown as Record<string, unknown>, [
    'isDisabled',
    'disabled',
    'formDisabled',
    'Disabled',
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

  React.useEffect(() => {
    setRequired(Boolean(isRequired));
  }, [isRequired]);

  React.useEffect(() => {
    const fromMode = isDisplayForm;
    const fromCtx = disabledFromCtx;
    const fromSubmitting = Boolean(submitting);
    const fromDisabledList = isListed(AllDisabledFields, displayName);
    const fromHiddenList = isListed(AllHiddenFields, displayName);

    setIsDisabled(fromMode || fromCtx || fromDisabledList || fromSubmitting);
    setIsHidden(fromHiddenList);

    const decision = evaluateFieldRules(props.id, {
      formMode: ctx.FormMode,
      formData: ctx.FormData,
      curUserInfo: ctx.curUserInfo,
      formConfigJson: ctx.formRules,
    });

    if (decision.isDisabled !== undefined) {
      setIsDisabled(decision.isDisabled || fromMode || fromCtx || fromDisabledList || fromSubmitting);
    }

    if (decision.isHidden !== undefined) {
      setIsHidden(decision.isHidden);
    } else {
      setIsHidden(false); // reset so it doesn't get stuck hidden
    }
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  /* ------------------------------ Load existing (Edit/View) ------------------------------ */
  React.useEffect((): void | (() => void) => {
    if (isNewMode) return;

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
            ? (attsRaw as unknown[])
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
                .filter((x): x is SPAttachment => !!x)
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
    })().catch(() => {
      /* no-op */
    });

    return (): void => {
      cancelled = true;
    };
  }, [isNewMode, context]);

  /* ------------------------------ Filtering, validation & committing ------------------------------ */

  // Convert selected files to Blobs and write to shared context.
  const commitWithBlob = React.useCallback(
    async (list: File[]): Promise<void> => {
      const blobItems = list.map((file) => ({
        name: file.name,
        content: fileToBlob(file), // explicit Blob conversion
      }));

      const payload: unknown =
        list.length === 0 ? undefined : multiple ? blobItems : blobItems[0];

      ctx.GlobalFormData(id, payload);
    },
    [ctx, id, multiple]
  );

  /* ------------------------------ Handlers ------------------------------ */

  const openPicker = (): void => {
    if (!isDisabled) inputRef.current?.click();
  };

  const onFilesPicked: React.ChangeEventHandler<HTMLInputElement> = async (e) => {
    const picked = Array.from(e.currentTarget.files ?? []);

    // Prepare dedupe baseline from current valid selection
    const existingNames = new Set(files.map((f) => normalizeName(f.name)));

    // Filter only the newly picked files (do not re-validate existing ones)
    const { validNew, warning } = filterNewFiles(
      multiple ? picked : picked.slice(0, 1),
      multiple ? existingNames : new Set<string>() // in single-file mode, always replace
    );

    // Ensure outgoing filenames are unique by base name, case-insensitively,
    // across both current selections and existing SharePoint attachments.
    const existingAllNames = new Set<string>([
      ...files.map((f) => normalizeName(f.name)),
      ...(spAttachments ?? []).map((a) => normalizeName(a.FileName)),
    ]);
    const uniqueNew = ensureUniqueFilenames(validNew, existingAllNames);

    // Compute the next set to show/commit
    const next = multiple ? files.concat(uniqueNew) : uniqueNew.slice(0, 1);

    // Validate combined size using ONLY next (post-filter & post-dedupe)
    const totalBytes = next.reduce((sum, f) => sum + (f?.size ?? 0), 0);
    if (totalBytes > TOTAL_LIMIT_BYTES) {
      setError(TOTAL_LIMIT_MSG); // local only
      ctx.GlobalErrorHandle(id, undefined); // do not set global for size
      if (inputRef.current) inputRef.current.value = '';
      return;
    }

    // Required rule after all filtering/dedupe
    const requiredMsg = required && next.length === 0 ? REQUIRED_MSG : '';

    // Local message: required takes precedence, else show any warnings (invalid/too-long/duplicates)
    const messageToShow = requiredMsg || warning;

    // Update state
    setFiles(next);
    setError(messageToShow);

    // Only propagate "required" to global; everything else remains local
    ctx.GlobalErrorHandle(id, requiredMsg ? requiredMsg : undefined);

    // Commit only what passed validation
    await commitWithBlob(next);

    if (inputRef.current) inputRef.current.value = '';
  };

  const removeAt = React.useCallback(
    async (idx: number): Promise<void> => {
      const next = files.filter((_, i) => i !== idx);

      const requiredMsg = required && next.length === 0 ? REQUIRED_MSG : '';

      const totalBytes = next.reduce((sum, f) => sum + (f?.size ?? 0), 0);
      const sizeMsg = totalBytes > TOTAL_LIMIT_BYTES ? TOTAL_LIMIT_MSG : '';

      const msg = requiredMsg || sizeMsg;

      setFiles(next);
      setError(msg);

      // Only set global for required; otherwise clear it
      ctx.GlobalErrorHandle(id, requiredMsg ? requiredMsg : undefined);

      if (!msg || requiredMsg === '') {
        await commitWithBlob(next);
      }
    },
    [files, required, ctx, id, commitWithBlob]
  );

  const clearAll = async (): Promise<void> => {
    const requiredMsg = required ? REQUIRED_MSG : '';

    setFiles([]);
    setError(requiredMsg);

    // Only set global for required; otherwise clear it
    ctx.GlobalErrorHandle(id, requiredMsg ? requiredMsg : undefined);

    await commitWithBlob([]);
    if (inputRef.current) inputRef.current.value = '';
  };

  /* ------------------------------ Delete existing file ------------------------------ */

  const deleteExistingAttachment = React.useCallback(
    async (fileName: string): Promise<void> => {
      if (!context) return;

      if (!window.confirm(`Are you sure you want to delete this file?\n\n${fileName}`)) return;

      const listTitle: string | undefined = (context as { list?: { title?: string } } | undefined)?.list?.title;
      const listGuid: string | undefined = (context as { list?: { id?: string } } | undefined)?.list?.id;
      const itemId: number | undefined = (context as { item?: { ID?: number } } | undefined)?.item?.ID;
      const baseUrl: string | undefined =
        (context as { pageContext?: { web?: { absoluteUrl?: string } } } | undefined)?.pageContext?.web?.absoluteUrl ??
        (typeof window !== 'undefined' ? window.location.origin : undefined);

      // SharePoint write/delete operations commonly require a request digest.
      const requestDigest: string | undefined =
        (
          context as {
            pageContext?: { legacyPageContext?: { formDigestValue?: string } };
          } | undefined
        )?.pageContext?.legacyPageContext?.formDigestValue;

      if (!baseUrl || !itemId || (!listGuid && !listTitle)) return;

      const encTitle = listTitle ? encodeURIComponent(listTitle) : '';
      const encFile = fileName.replace(/'/g, "''");
      const idStr = encodeURIComponent(String(itemId));

      const urls: string[] = [];
      if (listGuid) {
        urls.push(
          `${baseUrl}/_api/web/lists(guid'${listGuid}')/items(${idStr})/AttachmentFiles/getByFileName('${encFile}')`
        );
      }
      if (listTitle) {
        urls.push(
          `${baseUrl}/_api/web/lists/getbytitle('${encTitle}')/items(${idStr})/AttachmentFiles/getByFileName('${encFile}')`
        );
      }

      setDeletingName(fileName);
      setLoadError('');
      let success = false;
      let lastErr: unknown = null;

      for (const spUrl of urls) {
        try {
          await getFetchAPI({
            spUrl,
            method: 'DELETE',
            headers: {
              'IF-MATCH': '*',
              Accept: 'application/json;odata=nometadata',
              ...(requestDigest ? { 'X-RequestDigest': requestDigest } : {}),
            },
          });
          success = true;
          break;
        } catch (e1) {
          try {
            await getFetchAPI({
              spUrl,
              method: 'POST',
              headers: {
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE',
                Accept: 'application/json;odata=nometadata',
                ...(requestDigest ? { 'X-RequestDigest': requestDigest } : {}),
              },
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
        setSpAttachments((prev) =>
          Array.isArray(prev)
            ? prev.filter((a) => normalizeName(a.FileName) !== normalizeName(fileName))
            : prev
        );
      } else {
        const msg = lastErr instanceof Error ? lastErr.message : 'Failed to delete the attachment.';
        setLoadError(msg);
      }
    },
    [context]
  );

  /* ------------------------------ Render ------------------------------ */

  if (isHidden) return <div hidden className="fieldClass" />;

  return (
    <div className="fieldClass">
      <Field
        label={displayName}
        required={required}
        validationMessage={
          error ? (
            <span style={{ color: 'red', fontWeight: 700 }}>{error}</span>
          ) : undefined
        }
        validationState={error ? 'error' : undefined}
      >
        {/* Existing attachments in Edit/View */}
        {!isNewMode && (
          <div style={{ marginBottom: 8 }}>
            {loadingSP && (
              <Text size={200}>
                <Spinner size="tiny" />&nbsp;Loading attachments…
              </Text>
            )}

            {!loadingSP && loadError && (
              <Text
                size={200}
                aria-live="polite"
                style={{ color: 'red', fontWeight: 700 }}
              >
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
                        onClick={() => {
                          deleteExistingAttachment(a.FileName).catch(() => {});
                        }}
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

        {/* Hidden picker */}
        <input
          id={id}
          name={displayName}
          ref={inputRef}
          type="file"
          multiple={multiple}
          style={{ display: 'none' }}
          onChange={onFilesPicked}
          disabled={isDisabled}
        />

        {/* Actions */}
        <div
          className={className}
          style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}
        >
          <Button appearance="primary" icon={<AttachRegular />} onClick={openPicker} disabled={isDisabled}>
            {files.length === 0
              ? multiple
                ? 'Choose files'
                : 'Choose file'
              : multiple
                ? 'Add more files'
                : 'Choose different file'}
          </Button>

          {files.length > 0 && (
            <Button appearance="secondary" onClick={clearAll} icon={<DismissRegular />} disabled={isDisabled}>
              Clear
            </Button>
          )}
        </div>

        {/* New selections */}
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
                  onClick={() => {
                    removeAt(i).catch(() => {});
                  }}
                  disabled={isDisabled}
                  aria-label={`Remove ${f.name}`}
                />
              </div>
            ))}
          </div>
        )}

        {description !== '' && (
          <div className="descriptionText" style={{ marginTop: 6 }}>
            {description}
          </div>
        )}

        {/* Permanent guidance text under description — same style/color as description */}
        <div className="descriptionText" style={{ marginTop: 6, lineHeight: 1.4 }}>
          <div>
            <strong>Filename rules:</strong> Up to {MAX_NAME_LEN} characters; letters, numbers, spaces,
            underscores, and dots only.
          </div>
          <div>
            <strong>Size limit:</strong> Combined attachments must not exceed 250&nbsp;MB.
          </div>
          <div>
            <strong>Heads-up:</strong> Files with disallowed characters, duplicates, or overly long names will
            be skipped.
          </div>
        </div>
      </Field>
    </div>
  );
}