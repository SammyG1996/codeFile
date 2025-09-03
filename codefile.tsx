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

  placeholder?: string;
  className?: string;

  /** NEW: optional helper text shown after the input */
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
    placeholder,
    className,
    // default so the short-circuit check is always safe
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

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(max, max);
    return '';
  }, [isRequired, min, max, isNumericString]);

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
      const sanitized = isNumber ? decimalSanitizer(initial) : initial;
      setLocalVal(sanitized);
      setError('');
      setTouched(false);
      GlobalFormData(id, sanitized);
    } else {
      const existing = FormData ? toStr((FormData as any)[id]) : '';
      const sanitized = isNumber ? decimalSanitizer(existing) : existing;
      setLocalVal(sanitized);
      setError('');
      setTouched(false);
      GlobalFormData(id, sanitized);
    }
    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [FormMode, starterValue, id, isNumber, decimalSanitizer]);

  // Selection helper for paste (TS-safe)
  const getSelection = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  // TEXT: trim pasted content to fit; show Field-style error if truncated
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

  // Local change
  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data) => {
    const raw = data.value ?? '';
    const next = isNumber ? decimalSanitizer(raw) : raw;
    setLocalVal(next);

    // TEXT: show length error at/over cap; clears when below
    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      setError(lengthMsg);
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

  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      /* size prop intentionally omitted */
    >
      <Input
        id={inputId}
        name={id}
        className={className}
        placeholder={placeholder}
        value={localVal}
        onChange={handleChange}
        onBlur={handleBlur}
        onPaste={handlePaste}
        disabled={isDisabled}

        // TEXT ONLY
        maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}

        // NUMBER ONLY
        type={isNumber ? 'number' : 'text'}
        inputMode={isNumber ? 'decimal' : undefined}
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}
        contentAfter={after}
      />

      {/* NEW: description below input (short-circuit, strict equality) */}
      {description !== '' && <div className="descriptionText">{description}</div>}
    </Field>
  );
}
