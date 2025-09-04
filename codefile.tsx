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

  /** NEW: decimal places policy for number fields */
  decimalPlaces?: 'automatic' | 'one' | 'two';

  placeholder?: string;
  className?: string;

  /** Optional helper text rendered under the input */
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

  const inputId = useId('input');

  // Controlled state
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Mirror flags (reactive to prop changes)
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  const isNumber = type === 'number';
  const toStr = (v: unknown) => (v === null || v === undefined ? '' : String(v));

  // Allow negatives only if boundaries allow
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  // NEW: compute decimal limit (null = unlimited)
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

    // Decimal place check (even if typing handler trimmed, keep this guard)
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

  // Prefill: New (8) vs Edit/View
  React.useEffect(() => {
    if (FormMode === 8) {
      const initial = starterValue !== undefined ? toStr(starterValue) : '';
      const sanitized0 = isNumber ? decimalSanitizer(initial) : initial;
      const { value: sanitized, trimmed } = isNumber ? enforceDecimalLimit(sanitized0) : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setError(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
      setTouched(false);
      GlobalFormData(id, sanitized);
    } else {
      const existing = FormData ? toStr((FormData as any)[id]) : '';
      const sanitized0 = isNumber ? decimalSanitizer(existing) : existing;
      const { value: sanitized, trimmed } = isNumber ? enforceDecimalLimit(sanitized0) : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setError(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
      setTouched(false);
      GlobalFormData(id, sanitized);
    }
    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [FormMode, starterValue, id, isNumber, decimalSanitizer, enforceDecimalLimit, decimalLimit]);

  // Selection helper for paste (TS-safe)
  const getSelection = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  // TEXT: trim pasted content to fit maxLength and show Field-style error if truncated
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

    // Compose projected string as if paste happens into current selection
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
    const { value: next, trimmed } = isNumber ? enforceDecimalLimit(sanitized0) : { value: sanitized0, trimmed: false };

    setLocalVal(next);

    // TEXT: show length error at/over cap; clears when below
    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      setError(lengthMsg);
      return;
    }

    // NUMBER: if we trimmed due to decimal limit, surface the error immediately
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

  // Native step attribute to match decimal policy (helps arrows/increment and mobile keyboards)
  const stepAttr = isNumber
    ? (decimalLimit === 1 ? '0.1' : decimalLimit === 2 ? '0.01' : 'any')
    : undefined;

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

      {/* Description under the input */}
      {description !== '' && <div className="descriptionText">{description}</div>}
    </Field>
  );
}
