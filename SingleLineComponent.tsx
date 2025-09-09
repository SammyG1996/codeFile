/**
 * SingleLineComponent.tsx
 *
 * Purpose:
 * - Single-line input that can behave as text or number.
 * - Local state for value + error; pushes value/error to DynamicFormContext on BLUR.
 * - Supports: required, disabled, maxLength (text), min/max (number), decimalPlaces,
 *   optional "%" suffix, helper description, and a "submitting" flag from props.
 *
 * Notes for reviewers:
 * - The "submitting" prop is a simple boolean that disables the field and is passed to <Field>.
 * - Prefill runs on mount and whenever "submitting" changes; no GlobalFormData writes happen there.
 * - Value is validated/sanitized on change; final commit is on blur (can be made live if needed).
 */

import * as React from 'react';
import { Field, Input, Text, useId } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

export interface SingleLineFieldProps {
  id: string;
  displayName: string;

  starterValue?: string | number;                 // initial value for New mode or fallback
  isRequired?: boolean;                           // marks the field required
  disabled?: boolean;                             // externally disables the field

  // TEXT ONLY
  maxLength?: number;                             // hard cap for character count

  // NUMBER ONLY
  type?: 'number';                                // set to 'number' to enable numeric mode
  min?: number;                                   // inclusive minimum
  max?: number;                                   // inclusive maximum
  contentAfter?: 'percentage';                    // renders "%" suffix if 'percentage'
  decimalPlaces?: 'automatic' | 'one' | 'two';    // decimal precision policy; default 'automatic'

  placeholder?: string;
  className?: string;
  description?: string;                           // helper text under the field

  submitting?: boolean;                           // when true, disables input and marks Field as submitting
}

/* Messages shown via <Field validationMessage> */
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

/* Small helper for strict null checks */
const isDefined = <T,>(v: T | null | undefined): v is T => v !== null && v !== undefined;

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

  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  /* Expose submitting to Field and to local disable logic */
  const isSubmitting = !!props.submitting;

  const inputId = useId('input');

  /* Local state for controlled input + errors */
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  /* Mirror "required" and "disabled" props; keep them reactive */
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  /* Derived flags for number behavior */
  const isNumber = type === 'number';
  const toStr = (v: unknown) => (v === null || v === undefined ? '' : String(v));
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  /* Decimal policy: null means "no explicit limit" (automatic) */
  const decimalLimit: 1 | 2 | null = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return null;
  }, [decimalPlaces]);

  /* Keep only digits, one leading '-', and a single '.'; trim extra '-' or '.' occurrences */
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

  /* Legacy integer-only sanitizer kept here for reference; not used anymore */
  // const digitsOnly = (s: string) => s.replace(/[^\d]/g, '');

  /* Pre-built length message for text mode */
  const lengthMsg = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

  /* Utility for counting fractional digits in a string number */
  const getFractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
  };

  /* Enforce decimalPlaces policy by trimming extra fractional digits if needed */
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

  /* Text validation: just required check here */
  const validateText = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired]);

  /* Allow ".5" or "0.5" and optional leading "-" if negatives allowed */
  const isNumericString = React.useCallback((val: string): boolean => {
    if (!val || val.trim().length === 0) return false;
    const re = allowNegative
      ? /^-?(?:\d+\.?\d*|\.\d+)$/
      : /^(?:\d+\.?\d*|\.\d+)$/;
    return re.test(val);
  }, [allowNegative]);

  /* Number validation: required, format, decimal places, min/max */
  const validateNumber = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    if (val.trim().length === 0) return '';
    if (!isNumericString(val)) return INVALID_NUM_MSG;

    if (decimalLimit !== null && getFractionDigits(val) > decimalLimit) {
      return decimalLimitMsg(decimalLimit);
    }

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  /* Choose the validator based on mode */
  const computeError = React.useCallback(
    (val: string) => (isNumber ? validateNumber(val) : validateText(val)),
    [isNumber, validateNumber, validateText]
  );

  /* Push value + error to outer form state (called on BLUR) */
  const commitValue = React.useCallback((val: string, err: string) => {
    GlobalErrorHandle(id, err ? err : null);
    GlobalFormData(id, val);
  }, [GlobalErrorHandle, GlobalFormData, id]);

  /**
   * Prefill:
   * - Runs on mount and whenever props.submitting changes.
   * - Reads initial value from starterValue (New) or FormData[id] (Edit).
   * - Applies number sanitization and decimal trimming for number mode.
   * - Does NOT write to GlobalFormData here.
   * - Clears existing error, and disables input if submitting is true.
   */
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
    } else {
      const existing = FormData ? toStr((FormData as any)[id]) : '';
      const sanitized0 = isNumber ? decimalSanitizer(existing) : existing;
      const { value: sanitized, trimmed } = isNumber
        ? enforceDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setError(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
      setTouched(false);
    }

    if (props.submitting === true) {
      setIsDisabled(true);
    }

    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.submitting]);

  /* Selection positions for paste handling */
  const getSelection = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  /* TEXT paste: trim to fit maxLength and surface the length error if we had to cut */
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

  /* NUMBER paste: sanitize, enforce decimal limit, surface limit error if we trimmed */
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

  /* onChange: keep local value in sync, raise immediate errors when useful */
  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data) => {
    const raw = data.value ?? '';
    const sanitized0 = isNumber ? decimalSanitizer(raw) : raw;
    const { value: next, trimmed } = isNumber
      ? enforceDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setLocalVal(next);

    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      setError(lengthMsg);
      return;
    }

    if (isNumber && trimmed && decimalLimit !== null) {
      setError(decimalLimitMsg(decimalLimit));
      return;
    }

    const nextErr = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');
    setError(nextErr);

    // If you want live commits instead of blur-only, uncomment:
    // commitValue(next, nextErr);
  };

  /* onBlur: final validation + commit to outer state */
  const handleBlur: React.FocusEventHandler<HTMLInputElement> = () => {
    setTouched(true);
    const err =
      (!isNumber && isDefined(maxLength) && localVal.length >= maxLength)
        ? lengthMsg
        : computeError(localVal);
    setError(err);
    commitValue(localVal, err);
  };

  /* Optional suffix when contentAfter === 'percentage' */
  const after = isNumber && contentAfter === 'percentage'
    ? <Text size={400} id={`${inputId}Per`}>%</Text>
    : undefined;

  const hasError = error !== '';

  /* Step attribute aligns browser UI with our decimal policy (e.g., iOS numeric keypad) */
  const stepAttr = isNumber
    ? (decimalLimit === 1 ? '0.1' : decimalLimit === 2 ? '0.01' : 'any')
    : undefined;

  /* Allow custom 'submitting' prop on <Field> by casting */
  const FieldAny = Field as unknown as React.ComponentType<any>;

  return (
    <FieldAny
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      submitting={isSubmitting}
      /* no size prop to keep default sizing */
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

        /* TEXT ONLY */
        maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}

        /* NUMBER ONLY */
        type={isNumber ? 'number' : 'text'}
        inputMode={isNumber ? 'decimal' : undefined}
        step={stepAttr}
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}
        contentAfter={after}
      />

      {/* Helper text below the input */}
      {description !== '' && <div className="descriptionText">{description}</div>}
    </FieldAny>
  );
}
