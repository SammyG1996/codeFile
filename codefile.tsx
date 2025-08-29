import * as React from 'react';
import { Field, Input, Text, useId } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

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
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const INVALID_NUM_MSG = 'Please enter valid numeric value!';
const rangeMsg = (min?: number, max?: number) =>
  min != null && max != null ? `Value must be between ${min} and ${max}.`
  : min != null ? `Value must be ≥ ${min}.`
  : max != null ? `Value must be ≤ ${max}.`
  : '';

export default function SingleLineComponent(props: SingleLineFieldProps): JSX.Element {
  const {
    id, displayName, starterValue,
    isRequired: requiredProp, disabled: disabledProp,
    maxLength, type, min, max, contentAfter, placeholder
  } = props;

  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  const inputId = useId('input');

  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  const isNumber = type === 'number';

  // Allow negative only if either bound is negative (implicit behavior)
  const allowNegative = (min != null && min < 0) || (max != null && max < 0);

  const toStr = (v: unknown) => (v == null ? '' : String(v));

  // --- DECIMAL SANITIZER (single '.'; '-' only at start and only if allowed)
  const decimalSanitizer = React.useCallback((s: string): string => {
    let out = s.replace(/[^0-9.-]/g, '');         // keep digits, dot, minus
    if (!allowNegative) out = out.replace(/-/g, ''); // strip '-' if not allowed
    else {
      // keep only a single leading '-'
      const neg = out.startsWith('-');
      out = (neg ? '-' : '') + out.slice(neg ? 1 : 0).replace(/-/g, '');
    }
    // keep only the first dot
    const i = out.indexOf('.');
    if (i !== -1) {
      out = out.slice(0, i + 1) + out.slice(i + 1).replace(/\./g, '');
    }
    return out;
  }, [allowNegative]);

  // Preserve integer sanitizer for text->number transition clarity (unused now)
  const digitsOnly = (s: string) => s.replace(/[^\d]/g, '');

  const lengthMsg = maxLength != null ? `Maximum length is ${maxLength} characters.` : '';

  // --- validation
  const validateText = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired]);

  // Accepts: "12", "12.", "12.3", ".5", "-12.3" (if negatives allowed)
  const isNumericString = React.useCallback((val: string): boolean => {
    if (!val || !val.trim()) return false;
    const re = allowNegative
      ? /^-?(?:\d+\.?\d*|\.\d+)$/
      : /^(?:\d+\.?\d*|\.\d+)$/;
    return re.test(val);
  }, [allowNegative]);

  const validateNumber = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    if (val.trim().length === 0) return '';
    if (!isNumericString(val)) return INVALID_NUM_MSG;

    const n = Number(val); // safe after regex check
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (min != null && n < min) return rangeMsg(min, max);
    if (max != null && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString]);

  const computeError = React.useCallback(
    (val: string): string => (isNumber ? validateNumber(val) : validateText(val)),
    [isNumber, validateNumber, validateText]
  );

  const commitValue = React.useCallback((val: string, err: string) => {
    GlobalErrorHandle(id, err ? err : null);
    GlobalFormData(id, val);
  }, [GlobalErrorHandle, GlobalFormData, id]);

  // --- prefill: New (8) vs Edit/View
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

  // TS-safe selection helper (used for paste logic on text fields)
  const getSelection = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  // Trim pasted text to fit and show length error if we truncated (TEXT only)
  const handlePaste: React.ClipboardEventHandler<HTMLInputElement> = (e) => {
    if (isNumber || maxLength == null) return;

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

    const next = isNumber
      ? decimalSanitizer(raw)          // allow decimals (and '-' if allowed by range)
      : raw;

    setLocalVal(next);

    // TEXT: show length error whenever at/over cap; clears once below cap
    if (!isNumber && maxLength != null && next.length >= maxLength) {
      setError(lengthMsg);
      return;
    }

    // numbers: live-validate; text: defer required until blur
    const nextErr = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');
    setError(nextErr);

    // commitValue(next, nextErr); // opt-in for live commit
  };

  // Blur commit
  const handleBlur: React.FocusEventHandler<HTMLInputElement> = () => {
    setTouched(true);
    const err =
      !isNumber && maxLength != null && localVal.length >= maxLength
        ? lengthMsg
        : computeError(localVal);
    setError(err);
    commitValue(localVal, err);
  };

  // Optional % suffix
  const after = isNumber && contentAfter === 'percentage'
    ? <Text size={400} id={`${inputId}Per`}>%</Text>
    : undefined;

  const hasError = !!error;

  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      size="medium"
    >
      <Input
        id={inputId}
        name={id}
        placeholder={placeholder}
        value={localVal}
        onChange={handleChange}
        onBlur={handleBlur}
        onPaste={handlePaste}
        disabled={isDisabled}
        // text-only
        maxLength={!isNumber ? maxLength : undefined}
        // number-only
        type={isNumber ? 'number' : 'text'}
        inputMode={isNumber ? 'decimal' : undefined} // mobile keyboards show decimal keypad
        min={isNumber && min != null ? min : undefined}
        max={isNumber && max != null ? max : undefined}
        contentAfter={after}
      />
    </Field>
  );
}
