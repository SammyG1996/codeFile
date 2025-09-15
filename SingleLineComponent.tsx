/**
 * SingleLineComponent.tsx
 *
 * Summary
 * - Single-line input using Fluent UI v9 (<Field> + <Input>).
 * - TEXT mode (default) or NUMBER mode (when type==='number').
 * - Validations (required, text maxLength, number format, min/max, decimalPlaces).
 * - Disabled = (FormMode===4) OR (context disabled flags) OR (AllDisabledFields) OR (submitting).
 * - Hidden  = (AllHiddenFields) — hides the wrapper <div>.
 *
 * Behavior
 * - No global writes on mount/prefill.
 * - No GlobalFormData while typing; it commits only on BLUR.
 * - GlobalErrorHandle is called only after the field is touched (blurred once).
 * - When committing to GlobalFormData:
 *     · TEXT: empty string → null
 *     · NUMBER: empty/invalid → null; otherwise a real number
 *
 * Example usage
 * // TEXT
 * <SingleLineComponent
 *   id="title"
 *   displayName="Title"
 *   starterValue="Initial value"
 *   isRequired
 *   maxLength={120}
 *   placeholder="Enter title"
 *   description="Short helper text"
 *   submitting={isSubmitting}
 * />
 *
 * // NUMBER
 * <SingleLineComponent
 *   id="discount"
 *   displayName="Discount"
 *   type="number"
 *   min={0}
 *   max={100}
 *   decimalPlaces="two"
 *   contentAfter="percentage"
 *   starterValue={12.5}
 *   placeholder="e.g. 12.5"
 *   description="0–100, up to 2 decimals"
 *   submitting={isSubmitting}
 * />
 */

import * as React from 'react';
import { Field, Input, Text, useId } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

/* ---------- Props ---------- */

export interface SingleLineFieldProps {
  id: string;
  displayName: string;

  starterValue?: string | number;
  isRequired?: boolean;

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

  submitting?: boolean; // used to compute disabled; not forwarded to <Field>
}

/* ---------- Helpers ---------- */

const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const INVALID_NUM_MSG = 'Please enter valid numeric value!';

const rangeMsg = (min?: number, max?: number): string =>
  (min !== undefined && max !== undefined)
    ? `Value must be between ${min} and ${max}.`
    : (min !== undefined)
      ? `Value must be ≥ ${min}.`
      : (max !== undefined)
        ? `Value must be ≤ ${max}.`
        : '';

const decimalLimitMsg = (n: 1 | 2): string =>
  `Maximum ${n} decimal place${n === 1 ? '' : 's'} allowed.`;

// We treat "defined" as "not undefined" (we avoid runtime null checks)
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

export default function SingleLineComponent(props: SingleLineFieldProps): JSX.Element {
  const {
    id,
    displayName,
    starterValue,
    isRequired: requiredProp,
    maxLength,
    type,
    min,
    max,
    contentAfter,
    decimalPlaces = 'automatic',
    placeholder,
    className,
    description = '',
    submitting,
  } = props;

  // Use the context as-is; do not re-declare its shape locally.
  const formCtx = React.useContext(DynamicFormContext);
  const ctx = formCtx as unknown as Record<string, unknown>;

  // Pull required pieces from context (with safe guards)
  const FormData = hasKey(ctx, 'FormData') ? getKey<Record<string, unknown>>(ctx, 'FormData') : undefined;
  const FormMode = hasKey(ctx, 'FormMode') ? getKey<number>(ctx, 'FormMode') : undefined;

  // These two must exist on the provider; we assert-read them
  const GlobalFormData = getKey<(id: string, value: unknown) => void>(ctx, 'GlobalFormData');
  const GlobalErrorHandle = getKey<(id: string, error: string | null) => void>(ctx, 'GlobalErrorHandle');

  const isDisplayForm = FormMode === 4;
  const isNumber = type === 'number';

  const disabledFromCtx = getCtxFlag(ctx, ['isDisabled', 'disabled', 'formDisabled', 'Disabled']);

  // Disabled/hidden lists (optional on context)
  const AllDisabledFields = hasKey(ctx, 'AllDisabledFields') ? ctx.AllDisabledFields : undefined;
  const AllHiddenFields = hasKey(ctx, 'AllHiddenFields') ? ctx.AllHiddenFields : undefined;

  // Controlled flags
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || !!submitting || isListed(AllDisabledFields, displayName)
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(isListed(AllHiddenFields, displayName));

  // Value & validation state
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const inputId = useId('input');

  /* ---------- number helpers ---------- */

  const valueToString = (v: unknown): string => (v === undefined ? '' : String(v));
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  const decimalLimit: 1 | 2 | undefined = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return undefined;
  }, [decimalPlaces]);

  const sanitizeDecimal = React.useCallback((s: string): string => {
    let out = s.replace(/[^0-9.-]/g, '');
    if (!allowNegative) out = out.replace(/-/g, '');
    else {
      const neg = out.startsWith('-');
      out = (neg ? '-' : '') + out.slice(neg ? 1 : 0).replace(/-/g, '');
    }
    const i = out.indexOf('.');
    if (i !== -1) out = out.slice(0, i + 1) + out.slice(i + 1).replace(/\./g, '');
    return out;
  }, [allowNegative]);

  const lengthMsg = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

  const fractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
    };

  const applyDecimalLimit = React.useCallback(
    (val: string): { value: string; trimmed: boolean } => {
      if (decimalLimit === undefined) return { value: val, trimmed: false };
      const dot = val.indexOf('.');
      if (dot === -1) return { value: val, trimmed: false };
      const whole = val.slice(0, dot + 1);
      const frac = val.slice(dot + 1);
      if (frac.length <= decimalLimit) return { value: val, trimmed: false };
      return { value: whole + frac.slice(0, decimalLimit), trimmed: true };
    },
    [decimalLimit]
  );

  /* ---------- validation ---------- */

  const validateText = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired]);

  const isNumericString = React.useCallback((val: string): boolean => {
    if (!val || val.trim().length === 0) return false;
    const re = allowNegative ? /^-?(?:\d+\.?\d*|\.\d+)$/ : /^(?:\d+\.?\d*|\.\d+)$/;
    return re.test(val);
  }, [allowNegative]);

  const validateNumber = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    if (val.trim().length === 0) return '';
    if (!isNumericString(val)) return INVALID_NUM_MSG;

    if (decimalLimit !== undefined && fractionDigits(val) > decimalLimit) {
      return decimalLimitMsg(decimalLimit);
    }

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  const validate = React.useCallback((val: string): string => (
    isNumber ? validateNumber(val) : validateText(val)
  ), [isNumber, validateNumber, validateText]);

  /* ---------- commit + error ---------- */

  // GlobalFormData: ONLY on blur; empty → null
  const commitValue = React.useCallback((val: string): void => {
    if (isNumber) {
      const t = val.trim();
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = val.trim();
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(id, out === '' ? null : out);
    }
  }, [GlobalFormData, id, isNumber]);

  // GlobalErrorHandle: only after first blur (touched)
  const pushErrorIfTouched = React.useCallback((msg: string): void => {
    setError(msg);
    if (touched) {
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalErrorHandle(id, msg === '' ? null : msg);
    }
  }, [GlobalErrorHandle, id, touched]);

  /* ---------- effects ---------- */

  React.useEffect((): void => {
    setIsRequired(!!requiredProp);
  }, [requiredProp]);

  // Disabled/Hidden recompute
  React.useEffect((): void => {
    const fromMode = isDisplayForm;
    const fromCtx = disabledFromCtx;
    const fromSubmitting = !!submitting;
    const fromDisabledList = isListed(AllDisabledFields, displayName);
    const fromHiddenList = isListed(AllHiddenFields, displayName);

    setIsDisabled(fromMode || fromCtx || fromDisabledList || fromSubmitting);
    setIsHidden(fromHiddenList);
  }, [isDisplayForm, disabledFromCtx, AllDisabledFields, AllHiddenFields, displayName, submitting]);

  // Prefill once on mount (New vs Edit). No global commits here.
  React.useEffect((): void => {
    if (FormMode === 8) {
      const initial = starterValue !== undefined ? valueToString(starterValue) : '';
      const sanitized0 = isNumber ? sanitizeDecimal(initial) : initial;
      const { value: sanitized, trimmed } = isNumber
        ? applyDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setTouched(false);
      pushErrorIfTouched(trimmed && decimalLimit !== undefined ? decimalLimitMsg(decimalLimit) : '');
    } else {
      const existingRaw: unknown = FormData ? (FormData as Record<string, unknown>)[id] : '';
      const existing = valueToString(existingRaw);
      const sanitized0 = isNumber ? sanitizeDecimal(existing) : existing;
      const { value: sanitized, trimmed } = isNumber
        ? applyDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setTouched(false);
      pushErrorIfTouched(trimmed && decimalLimit !== undefined ? decimalLimitMsg(decimalLimit) : '');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // initialize once

  /* ---------- handlers ---------- */

  const getSelectionRange = (el: HTMLInputElement): { start: number; end: number } => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  const handleTextPaste: React.ClipboardEventHandler<HTMLInputElement> = (e): void => {
    if (isNumber || !isDefined(maxLength)) return;
    const input = e.currentTarget;
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;

    const { start, end } = getSelectionRange(input);
    const currentLen = input.value.length;
    const replacing = end - start;
    const spaceLeft = maxLength - (currentLen - replacing);

    if (spaceLeft <= 0) {
      e.preventDefault();
      pushErrorIfTouched(lengthMsg);
      return;
    }

    if (pasteText.length > spaceLeft) {
      e.preventDefault();
      const insert = pasteText.slice(0, Math.max(0, spaceLeft));
      const nextValue = input.value.slice(0, start) + insert + input.value.slice(end);
      setLocalVal(nextValue); // local only
      pushErrorIfTouched(lengthMsg);
    }
  };

  const handleNumberPaste: React.ClipboardEventHandler<HTMLInputElement> = (e): void => {
    if (!isNumber) return;
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;

    const input = e.currentTarget;
    const { start, end } = getSelectionRange(input);
    const projected = input.value.slice(0, start) + pasteText + input.value.slice(end);
    const sanitized0 = sanitizeDecimal(projected);
    const { value: limited, trimmed } = applyDecimalLimit(sanitized0);
    if (trimmed && decimalLimit !== undefined) {
      e.preventDefault();
      setLocalVal(limited);
      pushErrorIfTouched(decimalLimitMsg(decimalLimit));
    }
  };

  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data): void => {
    const raw = data.value ?? '';
    const sanitized0 = isNumber ? sanitizeDecimal(raw) : raw;
    const { value: next, trimmed } = isNumber
      ? applyDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setLocalVal(next); // local only

    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      pushErrorIfTouched(lengthMsg);
      return;
    }
    if (isNumber && trimmed && decimalLimit !== undefined) {
      pushErrorIfTouched(decimalLimitMsg(decimalLimit));
      return;
    }

    const currentError = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');
    pushErrorIfTouched(currentError);
  };

  const handleBlur: React.FocusEventHandler<HTMLInputElement> = (): void => {
    setTouched(true);
    const finalError =
      (!isNumber && isDefined(maxLength) && localVal.length >= maxLength)
        ? lengthMsg
        : validate(localVal);
    // update local + global error
    // eslint-disable-next-line @rushstack/no-new-null
    setError(finalError); GlobalErrorHandle(id, finalError === '' ? null : finalError);
    // single place we push to GlobalFormData
    commitValue(localVal);
  };

  /* ---------- render ---------- */

  const after = isNumber && contentAfter === 'percentage'
    ? <Text size={400} id={`${inputId}Per`}>%</Text>
    : undefined;

  return (
    <div hidden={isHidden} className="fieldClass">
      <Field
        label={displayName}
        required={isRequired}
        validationMessage={error !== '' ? error : undefined}
        validationState={error !== '' ? 'error' : undefined}
      >
        <Input
          id={inputId}
          name={id}
          className={className}
          placeholder={placeholder}
          value={localVal}
          onChange={handleChange}
          onBlur={handleBlur}
          onPaste={isNumber ? handleNumberPaste : handleTextPaste}
          disabled={isDisabled}
          /* TEXT ONLY */
          maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}
          /* NUMBER ONLY */
          type={isNumber ? 'number' : 'text'}
          inputMode={isNumber ? 'decimal' : undefined}
          step="any"
          min={isNumber && isDefined(min) ? min : undefined}
          max={isNumber && isDefined(max) ? max : undefined}
          contentAfter={after}
        />
        {description !== '' && <div className="descriptionText">{description}</div>}
      </Field>
    </div>
  );
}
