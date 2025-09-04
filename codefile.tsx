/**
 * SingleLineComponent.tsx
 *
 * USAGE
 * -----
 * TEXT mode (default):
 *   <SingleLineComponent
 *     id="title"                     // REQUIRED
 *     displayName="Title"            // REQUIRED
 *     maxLength={120}                // OPTIONAL
 *     isRequired={true}              // OPTIONAL
 *     disabled={false}               // OPTIONAL (will be overridden by AllDisabledFields / perms)
 *     starterValue="Prefilled text"  // OPTIONAL (used in New mode)
 *     placeholder="Enter title"      // OPTIONAL
 *     description="Shown under input as helper text" // OPTIONAL
 *     className="w-full"             // OPTIONAL
 *   />
 *
 * NUMBER mode:
 *   <SingleLineComponent
 *     id="discount"                  // REQUIRED
 *     displayName="Discount"         // REQUIRED
 *     type="number"                  // REQUIRED for number mode
 *     min={0}                        // OPTIONAL (inclusive)
 *     max={100}                      // OPTIONAL (inclusive)
 *     decimalPlaces="two"            // OPTIONAL: 'automatic' | 'one' | 'two' (default 'automatic')
 *     contentAfter="percentage"      // OPTIONAL: renders '%' suffix
 *     isRequired={true}              // OPTIONAL
 *     disabled={false}               // OPTIONAL (will be overridden by AllDisabledFields / perms)
 *     starterValue={12.5}            // OPTIONAL (used in New mode)
 *     placeholder="e.g. 12.5"        // OPTIONAL
 *     description="0 - 100, up to 2 decimals" // OPTIONAL
 *     className="w-48"               // OPTIONAL
 *   />
 *
 * NOTES
 * - Prefill (value seed + error clear) runs ONCE on mount.
 * - Permissions are recomputed reactively when context values change.
 * - Number mode supports decimals and (optionally) negatives (if min/max allow).
 * - decimalPlaces trims extra fraction digits while typing/pasting and surfaces a Field error.
 * - Old integer-only sanitizer is kept commented for reference.
 * - Hiding logic for `AllHiddenFields` / `visibleOnly` is included but COMMENTED OUT (enable when ready).
 */

import * as React from 'react';
import { Field, Input, Text, useId } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

/** Props */
export interface SingleLineFieldProps {
  id: string;
  displayName: string;

  starterValue?: string | number;
  isRequired?: boolean;
  disabled?: boolean;

  // TEXT ONLY
  maxLength?: number;

  // NUMBER ONLY
  type?: 'number';
  min?: number;
  max?: number;
  contentAfter?: 'percentage';
  decimalPlaces?: 'automatic' | 'one' | 'two';

  placeholder?: string;
  className?: string;
  description?: string;
}

/** Messages */
const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const INVALID_NUM_MSG = 'Please enter valid numeric value!';
const rangeMsg = (min?: number, max?: number) =>
  (min !== null && min !== undefined) && (max !== null && max !== undefined)
    ? `Value must be between ${min} and ${max}.`
    : (min !== null && min !== undefined)
      ? `Value must be ≥ ${min}.`
      : (max !== null && max !== undefined)
        ? `Value must be ≤ ${max}.`
        : '';
const decimalLimitMsg = (n: 1 | 2) =>
  `Maximum ${n} decimal place${n === 1 ? '' : 's'} allowed.`;

/** TS helper for strict null checks */
const isDefined = <T,>(v: T | null | undefined): v is T => v !== null && v !== undefined;

/** String normalizer for case/spacing-insensitive comparisons */
const norm = (s: unknown): string => String(s ?? '').trim().toLowerCase();

export default function SingleLineComponent(props: SingleLineFieldProps): JSX.Element {
  const {
    id,
    displayName,
    starterValue,
    isRequired: requiredProp,
    disabled: disabledProp,
    maxLength,
    type,
    min,
    max,
    contentAfter,
    decimalPlaces = 'automatic',
    placeholder,
    className,
    description = '',
  } = props;

  const {
    FormData,
    GlobalFormData,
    FormMode,
    GlobalErrorHandle,
    // NEW contexts used for permissions / visibility
    AllDisabledFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
  } = React.useContext(DynamicFormContext);

  const inputId = useId('input');

  // Controlled state
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Mirror flags (reactive to prop changes)
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  // Optional: keep a local hidden flag (not applied yet; left here for future use)
  // const [isHidden, setIsHidden] = React.useState<boolean>(false);

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  const isNumber = type === 'number';
  const toStr = (v: unknown) => (v === null || v === undefined ? '' : String(v));

  // Allow negatives only if boundaries allow
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  // Compute decimal limit (null = unlimited)
  const decimalLimit: 1 | 2 | null = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return null; // 'automatic'
  }, [decimalPlaces]);

  // DECIMAL sanitizer: one leading '-' (if allowed) + single '.'
  const decimalSanitizer = React.useCallback((s: string): string => {
    let out = s.replace(/[^0-9.-]/g, '');
    if (!allowNegative) {
      out = out.replace(/-/g, '');
    } else {
      const neg = out.startsWith('-');
      out = (neg ? '-' : '') + out.slice(neg ? 1 : 0).replace(/-/g, '');
    }
    const i = out.indexOf('.');
    if (i !== -1) {
      out = out.slice(0, i + 1) + out.slice(i + 1).replace(/\./g, '');
    }
    return out;
  }, [allowNegative]);

  // UNUSED (kept for reference; integer-only sanitizer from earlier versions)
  // const digitsOnly = (s: string) => s.replace(/[^\d]/g, '');

  const lengthMsg = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

  // Helpers for decimal limits
  const getFractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
  };

  const enforceDecimalLimit = React.useCallback(
    (val: string): { value: string; trimmed: boolean } => {
      if (decimalLimit === null) return { value: val, trimmed: false };
      const dot = val.indexOf('.');
      if (dot === -1) return { value: val, trimmed: false };
      const whole = val.slice(0, dot + 1);
      const frac = val.slice(dot + 1);
      if (frac.length <= decimalLimit) return { value: val, trimmed: false };
      return { value: whole + frac.slice(0, decimalLimit), trimmed: true };
    },
    [decimalLimit]
  );

  // Validation
  const validateText = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired]);

  // Accept: "12", "12.", "0.5", ".75", "-3.2" (if negatives allowed)
  const isNumericString = React.useCallback((val: string): boolean => {
    if (!val || val.trim().length === 0) return false;
    const re = allowNegative
      ? /^-?(?:\d+\.?\d*|\.\d+)$/
      : /^(?:\d+\.?\d*|\.\d+)$/;
    return re.test(val);
  }, [allowNegative]);

  const validateNumber = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    if (val.trim().length === 0) return '';
    if (!isNumericString(val)) return INVALID_NUM_MSG;

    // Decimal place check (guard even if typing handler trimmed)
    if (decimalLimit !== null && getFractionDigits(val) > decimalLimit) {
      return decimalLimitMsg(decimalLimit);
    }

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  const computeError = React.useCallback(
    (val: string) => (isNumber ? validateNumber(val) : validateText(val)),
    [isNumber, validateNumber, validateText]
  );

  const commitValue = React.useCallback((val: string, err: string) => {
    GlobalErrorHandle(id, err ? err : null);
    GlobalFormData(id, val);
  }, [GlobalErrorHandle, GlobalFormData, id]);

  // Prefill ONCE on mount only (requested earlier)
  React.useEffect(() => {
    if (FormMode === 8) {
      const initial = starterValue !== undefined ? toStr(starterValue) : '';
      const sanitized0 = isNumber ? decimalSanitizer(initial) : initial;
      const { value: sanitized, trimmed } = isNumber
        ? enforceDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setError(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
      setTouched(false);
      GlobalFormData(id, sanitized);
    } else {
      const existing = FormData ? toStr((FormData as any)[id]) : '';
      const sanitized0 = isNumber ? decimalSanitizer(existing) : existing;
      const { value: sanitized, trimmed } = isNumber
        ? enforceDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setError(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
      setTouched(false);
      GlobalFormData(id, sanitized);
    }
    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // run once

  /**
   * PERMISSIONS EFFECT (reactive):
   * - AllDisabledFields → always disable (highest priority).
   * - userBasedPerms:
   *    - Find groupName(s); if current user is in that SP group:
   *        - editableOnly contains this displayName → enabled; else disabled.
   *        - visibleOnly contains this displayName → visible; else hidden (NOT APPLIED YET; see comments below).
   *    - If user not in group → disabled.
   * - AllHiddenFields (NOT APPLIED YET; see comments below).
   */
  React.useEffect(() => {
    // Normalize component field name once
    const fieldNameNorm = norm(displayName);

    // 1) AllDisabledFields (highest precedence)
    const disabledByAllDisabled =
      Array.isArray(AllDisabledFields) &&
      AllDisabledFields.some((n: unknown) => norm(n) === fieldNameNorm);

    if (disabledByAllDisabled) {
      setIsDisabled(true);
      return; // no further checks; precedence applies
    }

    // 2) Determine user membership in the configured group(s)
    const groupsArray: any[] = Array.isArray(curUserInfo?.SPGroups) ? curUserInfo.SPGroups : [];
    const userGroupTitles = new Set(
      groupsArray
        .map((g) => norm((g && (g.title ?? g.Title)) as unknown as string))
        .filter((s) => s !== '')
    );

    // userBasedPerms expected shape per screenshot
    const entries: any[] = Array.isArray(userBasedPerms?.groupBased)
      ? userBasedPerms.groupBased
      : [];

    // Default policy if nothing matches: disabled
    let disabledByPerms = true;
    // let hiddenByPerms = false; // default visible for now

    for (const entry of entries) {
      const src = entry?.groupSource;
      if (!src) continue;

      const groupName: string = src.groupName as string;
      const editableOnly: unknown[] = Array.isArray(src.editableOnly) ? src.editableOnly : [];
      const visibleOnly: unknown[] = Array.isArray(src.visibleOnly) ? src.visibleOnly : [];

      // Is current user in this group?
      const inGroup = userGroupTitles.has(norm(groupName));

      if (!inGroup) {
        // If user is not in this group, this rule implies disabled; continue scanning other rules.
        // disabledByPerms remains true (disabled) unless another matching group enables it.
        continue;
      }

      // We have a group match; inspect field permissions
      const editable = editableOnly.some((n) => norm(n) === fieldNameNorm);
      const visible = visibleOnly.some((n) => norm(n) === fieldNameNorm);

      // EDITABILITY: if any matching group lists the field in editableOnly → enable
      if (editable) {
        disabledByPerms = false;
      } else {
        // keep disabled unless another matching group sets editable
        disabledByPerms = disabledByPerms && true;
      }

      // VISIBILITY (NOT APPLIED YET):
      // If any matching group lists the field in visibleOnly → visible; otherwise hidden.
      // When you're ready to enable this logic, un-comment the next lines
      // and wire the result to state and render.
      //
      // if (visible) {
      //   hiddenByPerms = false;
      // } else {
      //   hiddenByPerms = true;
      // }
    }

    // 3) AllHiddenFields (NOT APPLIED YET). Example of how it would work:
    // const hiddenByAllHidden =
    //   Array.isArray(AllHiddenFields) &&
    //   AllHiddenFields.some((n: unknown) => norm(n) === fieldNameNorm);
    //
    // const finalHidden = hiddenByAllHidden || hiddenByPerms;
    // setIsHidden(finalHidden);

    // Combine with the prop-level disabled flag
    const finalDisabled = !!disabledProp || disabledByPerms;
    setIsDisabled(finalDisabled);
  }, [
    displayName,
    disabledProp,
    AllDisabledFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
  ]);

  // Selection helper for paste (TS-safe)
  const getSelection = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  // TEXT: trim pasted content to fit maxLength and show error if truncated
  const handlePaste: React.ClipboardEventHandler<HTMLInputElement> = (e) => {
    if (isNumber || !isDefined(maxLength)) return;

    const input = e.currentTarget;
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;

    const { start, end } = getSelection(input);
    const currentLen = input.value.length;
    const replacing = end - start;
    const spaceLeft = maxLength - (currentLen - replacing);

    if (spaceLeft <= 0) {
      e.preventDefault();
      setError(lengthMsg);
      return;
    }

    if (pasteText.length > spaceLeft) {
      e.preventDefault();
      const insert = pasteText.slice(0, Math.max(0, spaceLeft));
      const nextValue = input.value.slice(0, start) + insert + input.value.slice(end);
      setLocalVal(nextValue);
      setError(lengthMsg);
    }
  };

  // NUMBER: On paste, enforce decimal limit and show limit error if trimming occurs
  const handleNumberPaste: React.ClipboardEventHandler<HTMLInputElement> = (e) => {
    if (!isNumber) return;
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;

    const input = e.currentTarget;
    const { start, end } = getSelection(input);
    const projected = input.value.slice(0, start) + pasteText + input.value.slice(end);
    const sanitized0 = decimalSanitizer(projected);
    const { value: limited, trimmed } = enforceDecimalLimit(sanitized0);
    if (trimmed && decimalLimit !== null) {
      e.preventDefault();
      setLocalVal(limited);
      setError(decimalLimitMsg(decimalLimit));
    }
  };

  // Local change
  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data) => {
    const raw = data.value ?? '';
    const sanitized0 = isNumber ? decimalSanitizer(raw) : raw;
    const { value: next, trimmed } = isNumber
      ? enforceDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setLocalVal(next);

    // TEXT: show length error at/over cap; clears when below
    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      setError(lengthMsg);
      return;
    }

    // NUMBER: if trimmed due to decimal limit, surface the error immediately
    if (isNumber && trimmed && decimalLimit !== null) {
      setError(decimalLimitMsg(decimalLimit));
      return;
    }

    // NUMBERS: live-validate; TEXT: defer required until blur unless touched
    const nextErr = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');
    setError(nextErr);

    // commitValue(next, nextErr); // uncomment for live commits
  };

  // Blur commit
  const handleBlur: React.FocusEventHandler<HTMLInputElement> = () => {
    setTouched(true);
    const err =
      (!isNumber && isDefined(maxLength) && localVal.length >= maxLength)
        ? lengthMsg
        : computeError(localVal);
    setError(err);
    commitValue(localVal, err);
  };

  // Optional % suffix
  const after = isNumber && contentAfter === 'percentage'
    ? <Text size={400} id={`${inputId}Per`}>%</Text>
    : undefined;

  const hasError = error !== '';

  // Align native spinner/keyboard with policy
  const stepAttr = isNumber
    ? (decimalLimit === 1 ? '0.1' : decimalLimit === 2 ? '0.01' : 'any')
    : undefined;

  // If/when you enable hiding, render nothing when hidden:
  // if (isHidden) return null;

  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      /* size intentionally omitted */
    >
      <Input
        id={inputId}
        name={id}
        className={className}
        placeholder={placeholder}
        value={localVal}
        onChange={handleChange}
        onBlur={handleBlur}
        onPaste={isNumber ? handleNumberPaste : handlePaste}
        disabled={isDisabled}

        // TEXT ONLY
        maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}

        // NUMBER ONLY
        type={isNumber ? 'number' : 'text'}
        inputMode={isNumber ? 'decimal' : undefined}
        step={stepAttr}
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}
        contentAfter={after}
      />

      {/* Description under the input (short-circuit + strict check) */}
      {description !== '' && <div className="descriptionText">{description}</div>}
    </Field>
  );
}
