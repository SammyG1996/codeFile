/**
 * SingleLineComponent.tsx
 *
 * Summary
 * - Single-line input using Fluent UI v9 (<Field> + <Input>).
 * - TEXT mode (default) or NUMBER mode (when type==='number').
 * - Local state for value + error.
 * - Validations:
 *   • required (text/number)
 *   • text maxLength (and error when user reaches the cap)
 *   • number format (".5", "12.", negatives if allowed)
 *   • min/max range (inclusive)
 *   • decimalPlaces: 'automatic' (no cap), 'one', or 'two'
 * - Disabled = (FormMode===4) OR (context disabled flags) OR (AllDisabledFields) OR (submitting).
 * - Hidden  = (AllHiddenFields) — hides the wrapper <div>.
 * - Live commit: whenever local value changes we ALSO call GlobalFormData
 *   (number -> real number or undefined; text -> string). Blur still validates & commits.
 *
 * Example usage
 * // TEXT mode
 * <SingleLineComponent
 *   id="title"
 *   displayName="Title"
 *   starterValue="Initial value"
 *   isRequired={true}
 *   maxLength={120}
 *   placeholder="Enter title"
 *   description="Short helper text"
 *   submitting={isSubmitting}
 * />
 *
 * // NUMBER mode
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

/* ---------- Types ---------- */

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

  submitting?: boolean; // used only to compute disabled; not forwarded to <Field>
}

interface DFMinimal {
  FormData?: Record<string, unknown>;
  FormMode?: number;
  GlobalFormData: (id: string, value: unknown) => void;
  // We avoid `null` in this component; provider can normalize undefined→null if needed.
  GlobalErrorHandle: (id: string, error?: string) => void;

  // Optional flags/lists that may or may not exist on the context:
  // isDisabled / disabled / formDisabled / Disabled
  // AllDisabledFields / AllHiddenFields
}

/* ---------- Constants & helpers ---------- */

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

// Treat "defined" as "not undefined" — we never use `null` in this component.
const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

/** Safely read a boolean-ish flag from context using several candidate keys. */
const getCtxFlag = (ctx: Record<string, unknown>, keys: string[]): boolean => {
  for (const k of keys) {
    if (Object.prototype.hasOwnProperty.call(ctx, k)) {
      return !!ctx[k];
    }
  }
  return false;
};

/** truthy conversion for object maps */
const toBool = (v: unknown): boolean => !!v;

/**
 * Membership over unknown list-like:
 * - string[]  -> exact (case-insensitive) name match
 * - Set       -> iterate and match
 * - string    -> comma-separated
 * - object    -> keys are names; value truthy enables the entry
 */
const isListed = (bag: unknown, name: string): boolean => {
  const needle = name.trim().toLowerCase();
  if (!bag) return false;

  if (Array.isArray(bag)) {
    return bag.some(v => String(v).trim().toLowerCase() === needle);
  }

  // Set-like (duck-typed)
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

  // Context (no `any` at call sites)
  const formCtx = React.useContext(DynamicFormContext) as unknown as DFMinimal & Record<string, unknown>;
  const { FormData, FormMode, GlobalFormData, GlobalErrorHandle } = formCtx;

  const isDisplayForm: boolean = FormMode === 4;
  const isNumber: boolean = type === 'number';

  // Context-level disabled flags
  const disabledFromCtx: boolean = getCtxFlag(formCtx, [
    'isDisabled',
    'disabled',
    'formDisabled',
    'Disabled',
  ]);

  // UI flags (controlled)
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || disabledFromCtx || !!submitting
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

  // Field state
  const [localVal, _setLocalVal] = React.useState<string>('');
  const [error, _setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const inputId: string = useId('input');

  /* ---------- number helpers ---------- */

  const valueToString = (v: unknown): string => (v === undefined ? '' : String(v));
  const allowNegative: boolean = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  const decimalLimit: 1 | 2 | undefined = React.useMemo<1 | 2 | undefined>(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return undefined;
  }, [decimalPlaces]);

  const sanitizeDecimal = React.useCallback((s: string): string => {
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

  const lengthMsg: string = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

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
    const re = allowNegative
      ? /^-?(?:\d+\.?\d*|\.\d+)$/
      : /^(?:\d+\.?\d*|\.\d+)$/;
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

  /* ---------- error sync ---------- */

  const setErrorBoth = React.useCallback((msg: string): void => {
    _setError(msg);
    // Clear when empty; pass undefined (provider can coerce to null if required).
    if (msg === '') GlobalErrorHandle(id, undefined);
    else GlobalErrorHandle(id, msg);
  }, [GlobalErrorHandle, id]);

  /* ---------- live commit helper (local + GlobalFormData) ---------- */

  const setValueBoth = React.useCallback((raw: string): void => {
    _setLocalVal(raw);
    if (isNumber) {
      const trimmed = raw.trim();
      if (trimmed === '') {
        GlobalFormData(id, undefined);
      } else {
        const numeric = Number(trimmed);
        GlobalFormData(id, Number.isNaN(numeric) ? undefined : numeric);
      }
    } else {
      GlobalFormData(id, raw);
    }
  }, [GlobalFormData, id, isNumber]);

  /* ---------- commit helper (also updates local) ---------- */

  const commitValue = React.useCallback((val: string): void => {
    setValueBoth(val);
  }, [setValueBoth]);

  /* ---------- effects ---------- */

  React.useEffect((): void => {
    setIsRequired(!!requiredProp);
  }, [requiredProp]);

  // Disabled/Hidden computation (supporting many shapes in context)
  React.useEffect((): void => {
    const fromMode: boolean = isDisplayForm;
    const fromCtx: boolean = disabledFromCtx;
    const fromSubmitting: boolean = !!submitting;

    const disabledList: unknown = (formCtx as Record<string, unknown>).AllDisabledFields;
    const hiddenList: unknown = (formCtx as Record<string, unknown>).AllHiddenFields;

    const fromDisabledList: boolean = isListed(disabledList, displayName);
    const fromHiddenList: boolean = isListed(hiddenList, displayName);

    setIsDisabled(fromMode || fromCtx || fromDisabledList || fromSubmitting);
    setIsHidden(fromHiddenList);
  }, [isDisplayForm, disabledFromCtx, formCtx, displayName, submitting]);

  // Prefill once on mount (New vs Edit)
  React.useEffect((): void => {
    if (FormMode === 8) {
      const initial: string = starterValue !== undefined ? valueToString(starterValue) : '';
      const sanitized0: string = isNumber ? sanitizeDecimal(initial) : initial;
      const limited = isNumber ? applyDecimalLimit(sanitized0) : { value: sanitized0, trimmed: false };
      setValueBoth(limited.value);
      setTouched(false);
      setErrorBoth(limited.trimmed && decimalLimit !== undefined ? decimalLimitMsg(decimalLimit) : '');
    } else {
      const existingRaw: unknown = FormData ? (FormData as Record<string, unknown>)[id] : '';
      const existing: string = valueToString(existingRaw);
      const sanitized0: string = isNumber ? sanitizeDecimal(existing) : existing;
      const limited = isNumber ? applyDecimalLimit(sanitized0) : { value: sanitized0, trimmed: false };
      setValueBoth(limited.value);
      setTouched(false);
      setErrorBoth(limited.trimmed && decimalLimit !== undefined ? decimalLimitMsg(decimalLimit) : '');
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
      setErrorBoth(lengthMsg);
      return;
    }

    if (pasteText.length > spaceLeft) {
      e.preventDefault();
      const insert = pasteText.slice(0, Math.max(0, spaceLeft));
      const nextValue = input.value.slice(0, start) + insert + input.value.slice(end);
      setValueBoth(nextValue); // live commit
      setErrorBoth(lengthMsg);
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
      setValueBoth(limited); // live commit
      setErrorBoth(decimalLimitMsg(decimalLimit));
    }
  };

  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data): void => {
    const raw = data.value ?? '';
    const sanitized0 = isNumber ? sanitizeDecimal(raw) : raw;
    const limited = isNumber ? applyDecimalLimit(sanitized0) : { value: sanitized0, trimmed: false };

    setValueBoth(limited.value); // live commit

    if (!isNumber && isDefined(maxLength) && limited.value.length >= maxLength) {
      setErrorBoth(lengthMsg);
      return;
    }

    if (isNumber && limited.trimmed && decimalLimit !== undefined) {
      setErrorBoth(decimalLimitMsg(decimalLimit));
      return;
    }

    const currentError = isNumber ? validateNumber(limited.value) : (touched ? validateText(limited.value) : '');
    setErrorBoth(currentError);
  };

  const handleBlur: React.FocusEventHandler<HTMLInputElement> = (): void => {
    setTouched(true);
    const finalError =
      (!isNumber && isDefined(maxLength) && localVal.length >= maxLength)
        ? lengthMsg
        : validate(localVal);
    setErrorBoth(finalError);
    commitValue(localVal); // also refreshes local + GlobalFormData
  };

  /* ---------- render ---------- */

  const after: React.ReactNode = isNumber && contentAfter === 'percentage'
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
