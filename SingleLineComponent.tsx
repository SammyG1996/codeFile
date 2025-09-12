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
 *   • decimalPlaces limit: 'automatic' (no cap), 'one', or 'two'
 * - Submitting and context flags control disabled; display form (mode 4) is always disabled.
 * - AllDisabledFields (context) → disable this field when it lists the field's displayName.
 * - AllHiddenFields (context) → hide the entire wrapper <div>.
 * - Live commit: every time local value changes we also call GlobalFormData
 *   (and convert to number/null in number mode). Blur still validates & commits.
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

  submitting?: boolean; // used only to compute disabled; not forwarded to Field
}

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

const isDefined = <T,>(v: T | null | undefined): v is T => v !== null && v !== undefined;

interface DFMinimal {
  FormData?: Record<string, unknown>;
  FormMode?: number;
  GlobalFormData: (id: string, value: unknown) => void;
  GlobalErrorHandle: (id: string, error: string | null) => void;
  AllDisabledFields?: string[];
  AllHiddenFields?: string[];
  // plus any other fields the app might put here…
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
    submitting,
  } = props;

  // Context
  const formCtx = React.useContext(DynamicFormContext) as unknown as (DFMinimal & Record<string, unknown>);
  const {
    FormData,
    FormMode,
    GlobalFormData,
    GlobalErrorHandle,
    AllDisabledFields = [],
    AllHiddenFields = [],
  } = formCtx;

  // Derived flags
  const isDisplayForm = FormMode === 4;
  const isNumber = type === 'number';
  const fromContextDisabled =
    !!(formCtx.isDisabled ?? formCtx.disabled ?? formCtx.formDisabled ?? (formCtx as any).Disabled); // eslint-disable-line @typescript-eslint/no-explicit-any

  // Controlled UI flags
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    isDisplayForm || fromContextDisabled || !!submitting
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

  // Field state
  const [localVal, _setLocalVal] = React.useState<string>('');
  const [error, _setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // IDs
  const inputId = useId('input');

  // —— helpers

  const valueToString = (v: unknown) => (v === null || v === undefined ? '' : String(v));
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  const decimalLimit: 1 | 2 | null = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return null;
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

  const lengthMsg = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

  const fractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
    };

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

    if (decimalLimit !== null && fractionDigits(val) > decimalLimit) {
      return decimalLimitMsg(decimalLimit);
    }

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  const validate = React.useCallback(
    (val: string) => (isNumber ? validateNumber(val) : validateText(val)),
    [isNumber, validateNumber, validateText]
  );

  // Keep Field error + Global error in sync
  const setErrorBoth = React.useCallback((msg: string) => {
    _setError(msg);
    GlobalErrorHandle(id, msg === '' ? null : msg);
  }, [GlobalErrorHandle, id]);

  // Live commit helper: updates local value AND writes to GlobalFormData
  const setValueBoth = React.useCallback((raw: string) => {
    _setLocalVal(raw);
    if (isNumber) {
      const trimmed = raw.trim();
      if (trimmed === '') {
        GlobalFormData(id, null);
      } else {
        const numeric = Number(trimmed);
        GlobalFormData(id, Number.isNaN(numeric) ? null : numeric);
      }
    } else {
      GlobalFormData(id, raw);
    }
  }, [GlobalFormData, id, isNumber]);

  // Commit helper (also updates local for consistency)
  const commitValue = React.useCallback((val: string) => {
    setValueBoth(val);
  }, [setValueBoth]);

  /* -------------------------
   * EFFECTS
   * -------------------------
   */

  // Mirror basic flags
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
  }, [requiredProp]);

  // Disable / Hide computation
  React.useEffect(() => {
    // Display mode (4) is always disabled
    const fromMode = isDisplayForm;

    // Global/context disabled flags
    const fromCtx = fromContextDisabled;

    // AllDisabledFields overrides (by displayName)
    const fromList = AllDisabledFields.includes(displayName);

    // submitting disables, but we don't pass it into Field
    const fromSubmitting = !!submitting;

    setIsDisabled(fromMode || fromCtx || fromList || fromSubmitting);

    // Hidden list (by displayName)
    setIsHidden(AllHiddenFields.includes(displayName));
  }, [
    isDisplayForm,
    fromContextDisabled,
    AllDisabledFields,
    AllHiddenFields,
    displayName,
    submitting,
  ]);

  // Prefill once on mount (New vs Edit)
  React.useEffect(() => {
    if (FormMode === 8) {
      const initial = starterValue !== undefined ? valueToString(starterValue) : '';
      const sanitized0 = isNumber ? sanitizeDecimal(initial) : initial;
      const { value: sanitized, trimmed } = isNumber
        ? applyDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setValueBoth(sanitized);
      setTouched(false);
      setErrorBoth(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
    } else {
      const existing = FormData ? valueToString((FormData as any)[id]) : ''; // eslint-disable-line @typescript-eslint/no-explicit-any
      const sanitized0 = isNumber ? sanitizeDecimal(existing) : existing;
      const { value: sanitized, trimmed } = isNumber
        ? applyDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setValueBoth(sanitized);
      setTouched(false);
      setErrorBoth(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // initialize once

  /* -------------------------
   * Handlers
   * -------------------------
   */

  const getSelectionRange = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

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
      setValueBoth(nextValue); // live commit
      setErrorBoth(lengthMsg);
    }
  };

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
      setValueBoth(limited); // live commit
      setErrorBoth(decimalLimitMsg(decimalLimit));
    }
  };

  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data) => {
    const raw = data.value ?? '';
    const sanitized0 = isNumber ? sanitizeDecimal(raw) : raw;
    const { value: next, trimmed } = isNumber
      ? applyDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setValueBoth(next); // live commit

    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      setErrorBoth(lengthMsg);
      return;
    }

    if (isNumber && trimmed && decimalLimit !== null) {
      setErrorBoth(decimalLimitMsg(decimalLimit));
      return;
    }

    const currentError = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');
    setErrorBoth(currentError);
  };

  const handleBlur: React.FocusEventHandler<HTMLInputElement> = () => {
    setTouched(true);
    const finalError =
      (!isNumber && isDefined(maxLength) && localVal.length >= maxLength)
        ? lengthMsg
        : validate(localVal);
    setErrorBoth(finalError);
    commitValue(localVal); // also refreshes local + GlobalFormData
  };

  /* -------------------------
   * Render
   * -------------------------
   */

  const after = isNumber && contentAfter === 'percentage'
    ? <Text size={400} id={`${inputId}Per`}>%</Text>
    : undefined;

  return (
    <div hidden={isHidden}>
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
