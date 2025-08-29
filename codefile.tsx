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
  const toStr = (v: unknown) => (v == null ? '' : String(v));
  const digitsOnly = (s: string) => s.replace(/[^\d]/g, ''); // integer-only

  const lengthMsg = maxLength != null ? `Maximum length is ${maxLength} characters.` : '';

  // --- validation
  const validateText = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired]);

  const validateNumber = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    if (val.trim().length === 0) return '';
    if (!/^\d+$/.test(val)) return INVALID_NUM_MSG;
    const n = Number(val);
    if (min != null && n < min) return rangeMsg(min, max);
    if (max != null && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max]);

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
      const sanitized = isNumber ? digitsOnly(initial) : initial;
      setLocalVal(sanitized);
      setError('');
      setTouched(false);
      GlobalFormData(id, sanitized);
    } else {
      const existing = FormData ? toStr((FormData as any)[id]) : '';
      const sanitized = isNumber ? digitsOnly(existing) : existing;
      setLocalVal(sanitized);
      setError('');
      setTouched(false);
      GlobalFormData(id, sanitized);
    }
    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [FormMode, starterValue, id, isNumber]);

  // --- selection helper (TS-safe)
  const getSelection = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  // Detect printable keys (ignore ctrl/cmd/alt, arrows, etc.)
  const isPrintable = (e: React.KeyboardEvent<HTMLInputElement>) =>
    !e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1;

  // Show the SAME Field error style when user attempts to exceed maxLength
  const handleKeyDown: React.ComponentProps<typeof Input>['onKeyDown'] = (e) => {
    if (isNumber || maxLength == null) return;

    const input = e.currentTarget;
    const { start, end } = getSelection(input);
    const replacing = end - start;
    const atCap = input.value.length >= maxLength;

    if (isPrintable(e) && replacing === 0 && atCap) {
      // native maxLength will block the char; we just surface the error message
      setError(lengthMsg);
    }
  };

  // Paste overflow handling (TEXT only) — trims to fit and shows error
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
    const next = isNumber ? digitsOnly(raw) : raw;

    setLocalVal(next);

    // numbers: live-validate; text: defer required until blur
    let nextErr = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');

    // If we were showing the length error, clear it as soon as we're back within the cap
    if (!isNumber && maxLength != null && next.length <= maxLength && error === lengthMsg) {
      nextErr = touched ? validateText(next) : '';
    }

    setError(nextErr);
    // commitValue(next, nextErr); // enable for live commit if desired
  };

  // Blur commit
  const handleBlur: React.FocusEventHandler<HTMLInputElement> = () => {
    setTouched(true);
    const err = computeError(localVal);
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
        onKeyDown={handleKeyDown}  // <-- shows Field-style error when typing at cap
        onPaste={handlePaste}      // <-- shows Field-style error when pasting beyond cap
        disabled={isDisabled}
        // text-only
        maxLength={!isNumber ? maxLength : undefined}
        // number-only
        type={isNumber ? 'number' : 'text'}
        min={isNumber && min != null ? min : undefined}
        max={isNumber && max != null ? max : undefined}
        contentAfter={after}
      />
    </Field>
  );
}
