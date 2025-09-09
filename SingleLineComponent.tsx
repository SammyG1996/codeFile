/**
 * SingleLineComponent.tsx
 *
 * Summary
 * - Single-line input using Fluent UI v9 (<Field> + <Input>).
 * - TEXT mode (default) or NUMBER mode (when type==='number').
 * - Local state for value + error. Commits value/error to DynamicFormContext on BLUR.
 * - Validations:
 *   • required (text/number)
 *   • text maxLength (and error when user reaches the cap)
 *   • number format (including ".5" and "-3.2" if negatives allowed)
 *   • min/max range (inclusive)
 *   • decimalPlaces limit: 'automatic' (no cap), 'one', or 'two'
 * - UI extras:
 *   • optional '%' suffix for number fields (contentAfter === 'percentage')
 *   • optional helper description under the input
 *   • submitting prop disables the input and is forwarded to <Field> (for styling/hooks)
 */

import * as React from 'react';
import { Field, Input, Text, useId } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

/** Component props (keep names consistent with the rest of your form system) */
export interface SingleLineFieldProps {
  id: string;                                   // key used in outer form state
  displayName: string;                          // label text

  starterValue?: string | number;               // seed value in New mode (FormMode === 8)
  isRequired?: boolean;                         // required flag
  disabled?: boolean;                           // external disable flag

  // TEXT ONLY
  maxLength?: number;                           // hard character cap

  // NUMBER ONLY
  type?: 'number';                              // enable number behavior when set
  min?: number;                                 // inclusive min
  max?: number;                                 // inclusive max
  contentAfter?: 'percentage';                  // render '%' suffix when set
  decimalPlaces?: 'automatic' | 'one' | 'two';  // fraction digit policy (default 'automatic')

  placeholder?: string;                         // input placeholder
  className?: string;                           // custom class for Input
  description?: string;                         // helper text rendered under the field

  submitting?: boolean;                         // when true, disable input + forward to <Field submitting>
}

/** Messages presented via <Field validationMessage> */
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

/** Narrow helper (strict null checks) */
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

  // form context (commit value/error on BLUR; prefill reads from FormData)
  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  // submitting flag from props (not from context)
  const isSubmitting = !!props.submitting;

  // generate a stable id for aria attributes, suffix id, etc.
  const inputId = useId('input');

  /* -------------------------
   * Local component state
   * -------------------------
   * - localVal: the controlled input value
   * - error: current validation message (shown by <Field>)
   * - touched: whether the user has left the field at least once (drives text required-on-blur)
   * - isRequired / isDisabled mirror props and react when props change
   */
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);

  // keep required/disabled in sync with props; also disable when submitting=true
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp || isSubmitting);
  }, [requiredProp, disabledProp, isSubmitting]);

  /* -------------------------
   * Mode & number helpers
   * -------------------------
   */
  const isNumber = type === 'number';
  const toStr = (v: unknown) => (v === null || v === undefined ? '' : String(v));

  // allow a leading '-' only if any bound permits negatives
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  // decimalPlaces policy: null means unlimited ("automatic")
  const decimalLimit: 1 | 2 | null = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return null;
  }, [decimalPlaces]);

  // constrain to digits + at most one leading '-' + a single '.'
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

  // legacy integer-only sanitizer (reference only; not used anymore)
  // const digitsOnly = (s: string) => s.replace(/[^\d]/g, '');

  // precompute the text length message so we can reuse it in handlers
  const lengthMsg = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

  // count fractional digits in a numeric string
  const getFractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
  };

  // enforce decimalPlaces by trimming extra fractional digits (returns the trimmed value + flag)
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

  /* -------------------------
   * Validation
   * -------------------------
   * - TEXT: only required check here (length is enforced by attribute and paste handling).
   * - NUMBER: format, decimalPlaces, min/max.
   */
  const validateText = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired]);

  // numeric grammar allows "12", "12.", "0.5", ".5" and optional leading '-' based on allowNegative
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

    if (decimalLimit !== null && getFractionDigits(val) > decimalLimit) {
      return decimalLimitMsg(decimalLimit);
    }

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  // pick the correct validator for the current mode
  const computeError = React.useCallback(
    (val: string) => (isNumber ? validateNumber(val) : validateText(val)),
    [isNumber, validateNumber, validateText]
  );

  /* -------------------------
   * Commit to outer state (BLUR)
   * -------------------------
   * - On blur we push both value and error to the form context.
   * - Prefill effect deliberately does NOT write to GlobalFormData (per your pattern).
   */
  const commitValue = React.useCallback((val: string, err: string) => {
    GlobalErrorHandle(id, err ? err : null);
    GlobalFormData(id, val);
  }, [GlobalErrorHandle, GlobalFormData, id]);

  /* -------------------------
   * Prefill effect (runs on mount and when props.submitting toggles)
   * -------------------------
   * - New mode (FormMode === 8): seed from starterValue.
   * - Edit mode: seed from FormData[id].
   * - Sanitize for number mode and enforce decimalPlaces early.
   * - Clear any existing error in context.
   * - Do NOT write GlobalFormData here.
   * - Also: if submitting is true, the separate effect above already disabled the field.
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

    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.submitting]); // re-run when submitting toggles, per requirement

  /* -------------------------
   * Handlers
   * -------------------------
   */

  // selection utility for paste logic (keeps TS happy with null positions)
  const getSelection = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  // TEXT paste: trim to fit maxLength and show length error if we cut content
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

  // NUMBER paste: sanitize and enforce decimal limit; if we trimmed, surface the limit error
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

  // onChange: keep local value in sync; raise immediate errors that help user correct input early
  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data) => {
    const raw = data.value ?? '';
    const sanitized0 = isNumber ? decimalSanitizer(raw) : raw;
    const { value: next, trimmed } = isNumber
      ? enforceDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setLocalVal(next);

    // TEXT: show error when user reaches the cap (they can't type past maxLength anyway)
    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      setError(lengthMsg);
      return;
    }

    // NUMBER: if we trimmed due to decimalPlaces, show that error immediately
    if (isNumber && trimmed && decimalLimit !== null) {
      setError(decimalLimitMsg(decimalLimit));
      return;
    }

    // otherwise, compute the current error
    const nextErr = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');
    setError(nextErr);

    // If you prefer live commits, uncomment:
    // commitValue(next, nextErr);
  };

  // onBlur: final validation + commit to outer form state
  const handleBlur: React.FocusEventHandler<HTMLInputElement> = () => {
    setTouched(true);
    const err =
      (!isNumber && isDefined(maxLength) && localVal.length >= maxLength)
        ? lengthMsg
        : computeError(localVal);
    setError(err);
    commitValue(localVal, err);
  };

  /* -------------------------
   * Render
   * -------------------------
   */

  // optional suffix element (kept small and semantic)
  const after = isNumber && contentAfter === 'percentage'
    ? <Text size={400} id={`${inputId}Per`}>%</Text>
    : undefined;

  const hasError = error !== '';

  // align the native step with our decimal policy (helps browsers & mobile keyboards)
  const stepAttr = isNumber
    ? (decimalLimit === 1 ? '0.1' : decimalLimit === 2 ? '0.01' : 'any')
    : undefined;

  // allow custom 'submitting' prop on <Field> by casting (not part of official Field props)
  const FieldAny = Field as unknown as React.ComponentType<any>;

  return (
    <FieldAny
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      submitting={isSubmitting}
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

      {/* helper/description under the input (short-circuit pattern) */}
      {description !== '' && <div className="descriptionText">{description}</div>}
    </FieldAny>
  );
}
