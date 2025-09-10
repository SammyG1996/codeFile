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
 *   • number format (".5", "12.", negatives if allowed)
 *   • min/max range (inclusive)
 *   • decimalPlaces limit: 'automatic' (no cap), 'one', or 'two'
 * - UI extras:
 *   • optional '%' suffix for number fields (contentAfter === 'percentage')
 *   • optional helper description under the input
 *   • submitting prop disables the input and is forwarded to <Field> for styling
 * - "disabled" is NOT a prop. It’s derived from DynamicFormContext.
 *
 * Example usage (TypeScript/JSX):
 *
 * // TEXT mode
 * <SingleLineComponent
 *   id="title"                     // string (required)
 *   displayName="Title"            // string (required)
 *   starterValue="Initial value"   // string | number (optional)
 *   isRequired={true}              // boolean (optional)
 *   maxLength={120}                // number (optional; TEXT only)
 *   placeholder="Enter title"      // string (optional)
 *   description="Short helper text under the field" // string (optional)
 *   submitting={isSubmitting}      // boolean (optional)
 * />
 *
 * // NUMBER mode
 * <SingleLineComponent
 *   id="discount"                  // string (required)
 *   displayName="Discount"         // string (required)
 *   type="number"                  // 'number' (required for number mode)
 *   min={0}                        // number (optional)
 *   max={100}                      // number (optional)
 *   decimalPlaces="two"            // 'automatic' | 'one' | 'two' (optional)
 *   contentAfter="percentage"      // shows '%' to the right (optional)
 *   starterValue={12.5}            // number (optional)
 *   placeholder="e.g. 12.5"        // string (optional)
 *   description="0–100, up to 2 decimals" // string (optional)
 *   submitting={isSubmitting}      // boolean (optional)
 * />
 * 
 */

import * as React from 'react';
import { Field, Input, Text, useId } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

/** Props */
export interface SingleLineFieldProps {
  id: string;                                   // key used in outer form state
  displayName: string;                          // label text

  starterValue?: string | number;               // seed value in New mode (FormMode === 8)
  isRequired?: boolean;                         // required flag

  // TEXT ONLY
  maxLength?: number;                           // hard character cap

  // NUMBER ONLY
  type?: 'number';                              // enable number behavior when set
  min?: number;                                 // inclusive min
  max?: number;                                 // inclusive max
  contentAfter?: 'percentage';                  // render '%' suffix when set
  decimalPlaces?: 'automatic' | 'one' | 'two';  // fraction digit policy (default 'automatic')

  placeholder?: string;                         // input placeholder
  className?: string;                           // custom class for <Input>
  description?: string;                         // helper text under the field

  submitting?: boolean;                         // when true, disable input + forward to <Field submitting>
}

/** Messages shown via <Field validationMessage> */
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

/** Strict-null helper */
const isDefined = <T,>(v: T | null | undefined): v is T => v !== null && v !== undefined;

/** Minimal context typing (only what we use directly) */
interface DFMinimal {
  FormData?: Record<string, any>;
  FormMode?: number;
  GlobalFormData: (id: string, value: any) => void;
  GlobalErrorHandle: (id: string, error: string | null) => void;
}

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
  } = props;

  // Full context value (also read a "disabled" flag from here)
  const formCtx = React.useContext(DynamicFormContext) as unknown as (DFMinimal & Record<string, any>);
  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } = formCtx;

  // "disabled" comes from context (support several possible keys)
  const disabledFromContext: boolean = !!(
    formCtx.isDisabled ?? formCtx.disabled ?? formCtx.formDisabled ?? formCtx.Disabled
  );

  const isSubmitting = !!props.submitting;              // submitting flag from props
  const inputId = useId('input');                       // stable id for aria/suffix

  /* -------------------------
   * Local state
   * -------------------------
   */
  const [localVal, setLocalVal] = React.useState<string>('');   // controlled input value
  const [error, _setError] = React.useState<string>('');        // current validation message
  const [touched, setTouched] = React.useState<boolean>(false); // used to defer text required check until blur

  // mirror required + computed disabled; keep reactive
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(disabledFromContext || isSubmitting);

  // keep <Field> error and GlobalErrorHandle fully in sync everywhere
  const setErrorBoth = React.useCallback((msg: string) => {
    _setError(msg);
    GlobalErrorHandle(id, msg === '' ? null : msg);
  }, [GlobalErrorHandle, id]);

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(disabledFromContext || isSubmitting); // effective disabled
  }, [requiredProp, disabledFromContext, isSubmitting]);

  /* -------------------------
   * Mode & number helpers
   * -------------------------
   */
  const isNumber = type === 'number';
  const valueToString = (v: unknown) => (v === null || v === undefined ? '' : String(v));

  // allow a leading '-' only if any bound permits negatives
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  // decimalPlaces policy: null means unlimited ("automatic")
  const decimalLimit: 1 | 2 | null = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return null;
  }, [decimalPlaces]);

  // constrain to digits + at most one leading '-' + a single '.'
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

  // Legacy integer-only sanitizer kept for reference (unused)
  // const digitsOnly = (s: string) => s.replace(/[^\d]/g, '');

  // precompute the text length message so we can reuse it in handlers
  const lengthMsg = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

  // count fractional digits in a numeric string
  const fractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
  };

  // enforce decimalPlaces by trimming extra fractional digits
  const applyDecimalLimit = React.useCallback(
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

    if (decimalLimit !== null && fractionDigits(val) > decimalLimit) {
      return decimalLimitMsg(decimalLimit);
    }

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  // Pick the correct validator for the current mode
  const validate = React.useCallback(
    (val: string) => (isNumber ? validateNumber(val) : validateText(val)),
    [isNumber, validateNumber, validateText]
  );

  /* -------------------------
   * Commit to outer state (on BLUR)
   * -------------------------
   * - If number mode, commit a numeric value (not a string).
   * - Empty numeric input commits null (avoids NaN and avoids committing a string).
   */
  const commitValue = React.useCallback((val: string) => {
    if (isNumber) {
      const trimmed = val.trim();
      if (trimmed === '') {
        GlobalFormData(id, null);
        return;
      }
      const numeric = Number(trimmed);
      GlobalFormData(id, Number.isNaN(numeric) ? null : numeric);
      return;
    }
    GlobalFormData(id, val);
  }, [GlobalFormData, id, isNumber]);

  /* -------------------------
   * Prefill (mount + when submitting toggles)
   * -------------------------
   * - New (FormMode === 8): seed from starterValue.
   * - Edit: seed from FormData[id].
   * - Sanitize/trim for number mode.
   * - Keep Field + Global error in sync via setErrorBoth('').
   * - No GlobalFormData write here.
   */
  React.useEffect(() => {
    if (FormMode === 8) {
      const initial = starterValue !== undefined ? valueToString(starterValue) : '';
      const sanitized0 = isNumber ? sanitizeDecimal(initial) : initial;
      const { value: sanitized, trimmed } = isNumber
        ? applyDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setTouched(false);
      setErrorBoth(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
    } else {
      const existing = FormData ? valueToString((FormData as any)[id]) : '';
      const sanitized0 = isNumber ? sanitizeDecimal(existing) : existing;
      const { value: sanitized, trimmed } = isNumber
        ? applyDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setTouched(false);
      setErrorBoth(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.submitting]);

  /* -------------------------
   * Handlers
   * -------------------------
   */

  // selection utility for paste logic (handles null selectionStart/End safely)
  const getSelectionRange = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  // TEXT paste: trim to fit maxLength and show length error if cut
  const handleTextPaste: React.ClipboardEventHandler<HTMLInputElement> = (e) => {
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
      setLocalVal(nextValue);
      setErrorBoth(lengthMsg);
    }
  };

  // NUMBER paste: sanitize and enforce decimal limit; if trimmed, surface the limit error
  const handleNumberPaste: React.ClipboardEventHandler<HTMLInputElement> = (e) => {
    if (!isNumber) return;
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;

    const input = e.currentTarget;
    const { start, end } = getSelectionRange(input);
    const projected = input.value.slice(0, start) + pasteText + input.value.slice(end);
    const sanitized0 = sanitizeDecimal(projected);
    const { value: limited, trimmed } = applyDecimalLimit(sanitized0);
    if (trimmed && decimalLimit !== null) {
      e.preventDefault();
      setLocalVal(limited);
      setErrorBoth(decimalLimitMsg(decimalLimit));
    }
  };

  // onChange: keep local value in sync; raise immediate, helpful errors
  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data) => {
    const raw = data.value ?? '';
    const sanitized0 = isNumber ? sanitizeDecimal(raw) : raw;
    const { value: next, trimmed } = isNumber
      ? applyDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setLocalVal(next);

    // TEXT: show error when user reaches the cap
    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      setErrorBoth(lengthMsg);
      return;
    }

    // NUMBER: if we trimmed due to decimalPlaces, show that error immediately
    if (isNumber && trimmed && decimalLimit !== null) {
      setErrorBoth(decimalLimitMsg(decimalLimit));
      return;
    }

    // otherwise compute the current error
    const currentError = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');
    setErrorBoth(currentError);

    // For live commit instead of blur-only, you could call:
    // commitValue(next);
  };

  // onBlur: final validation + commit to outer form state
  const handleBlur: React.FocusEventHandler<HTMLInputElement> = () => {
    setTouched(true);
    const finalError =
      (!isNumber && isDefined(maxLength) && localVal.length >= maxLength)
        ? lengthMsg
        : validate(localVal);
    setErrorBoth(finalError);
    commitValue(localVal); // commits number (as number) or text (as string)
  };

  /* -------------------------
   * Render
   * -------------------------
   */

  // optional '%' suffix
  const after = isNumber && contentAfter === 'percentage'
    ? <Text size={400} id={`${inputId}Per`}>%</Text>
    : undefined;

  // Allow custom 'submitting' prop on <Field> by casting (not an official v9 prop)
  const FieldAny = Field as unknown as React.ComponentType<any>;

  return (
    <FieldAny
      label={displayName}
      required={isRequired}
      validationMessage={error !== '' ? error : undefined}
      validationState={error !== '' ? 'error' : undefined}
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
        onPaste={isNumber ? handleNumberPaste : handleTextPaste}
        disabled={isDisabled}
        /* TEXT ONLY */
        maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}
        /* NUMBER ONLY */
        type={isNumber ? 'number' : 'text'}
        inputMode={isNumber ? 'decimal' : undefined}
        step="any"  // always allow any precision; our validation enforces decimalPlaces
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}
        contentAfter={after}
      />

      {/* helper/description under the input */}
      {description !== '' && <div className="descriptionText">{description}</div>}
    </FieldAny>
  );
}