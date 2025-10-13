/**
 * SingleLineComponent.tsx
 *
 * What this component does
 * ------------------------
 * Reusable single-line input (<Field> + <Input>) that supports:
 *  • TEXT (default)
 *  • NUMBER (min/max, decimalPlaces, paste sanitization, optional “%” after input)
 *  • FILE display (shows file base name in the field, extension in contentAfter)
 *
 * Integrates with DynamicFormContext to:
 *  • Prefill from FormData / FormMode
 *  • Commit on blur (and on submit if focused)
 *  • Report validation via GlobalErrorHandle
 *  • Expose a ref via GlobalRefs
 *  • Apply centralized rules via formFieldsSetup (aligned with ComboBox pattern)
 *
 * Example
 * -------
 * <SingleLineComponent id="title" displayName="Title" maxLength={120} isRequired />
 * <SingleLineComponent id="discount" displayName="Discount" type="number" min={0} max={100} decimalPlaces="two" contentAfter="percentage" />
 * <SingleLineComponent id="docName" displayName="Document Name" type="file" starterValue="Proposal_v3.docx" />
 */

import * as React from 'react';
import { Field, Input, Text } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import { formFieldsSetup, FormFieldsProps } from '../Utils/formFieldBased';

const DEBUG = true;

/* ───────────────────────────── Props ──────────────────────────── */
export interface SingleLineFieldProps {
  id: string;
  displayName: string;

  starterValue?: string | number;
  isRequired?: boolean;

  // TEXT / FILE
  maxLength?: number;

  // Type selector (TEXT is default)
  type?: 'text' | 'number' | 'file';

  // NUMBER
  min?: number;
  max?: number;
  decimalPlaces?: 'automatic' | 'one' | 'two';
  contentAfter?: 'percentage';

  placeholder?: string;
  className?: string;
  description?: string;

  submitting?: boolean;
}

/* ───────────────────── Helpers & messages ─────────────────────── */

const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const INVALID_NUM_MSG = 'Please enter valid numeric value!';
const decimalLimitMsg = (n: 1 | 2) => `Maximum ${n} decimal place${n === 1 ? '' : 's'} allowed.`;
const rangeMsg = (min?: number, max?: number) =>
  (min !== undefined && max !== undefined)
    ? `Value must be between ${min} and ${max}.`
    : (min !== undefined)
      ? `Value must be ≥ ${min}.`
      : (max !== undefined)
        ? `Value must be ≤ ${max}.`
        : '';

const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

/** Split filename into base + extension (ext includes dot, e.g. ".docx"). */
function splitExt(name: string): { base: string; ext: string } {
  const i = name.lastIndexOf('.');
  if (i <= 0 || i === name.length - 1) return { base: name, ext: '' };
  return { base: name.slice(0, i), ext: name.slice(i) };
}

/* ─────────────────────────── Component ────────────────────────── */

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
    decimalPlaces = 'automatic',
    contentAfter,
    placeholder,
    className,
    description = '',
    submitting,
  } = props;

  /* ----- Context (match ComboBox naming; provide fallbacks) ----- */
  const ctx = React.useContext(DynamicFormContext) as Record<string, unknown>;
  const {
    FormData,
    GlobalFormData,
    FormMode,
    GlobalErrorHandle,
    GlobalRefs,

    // ComboBox provider often exposes this as "AllDisableFields" (no 'd')
    AllDisableFields,
    AllDisabledFields,

    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
  } = (ctx as any) ?? {};

  // Normalize for the rest of the file
  const AllDisabledFieldsNorm = (AllDisableFields ?? AllDisabledFields) as unknown;
  const AllHiddenFieldsNorm = AllHiddenFields as unknown;

  // Mount-only debug of context keys
  React.useEffect(() => {
    if (!DEBUG) return;
    // eslint-disable-next-line no-console
    console.log('[SingleLineComponent] context keys:', Object.keys(ctx || {}));
    // eslint-disable-next-line no-console
    console.log('[SingleLineComponent] disabled list presence', {
      hasAllDisableFields: AllDisableFields !== undefined,
      hasAllDisabledFields: AllDisabledFields !== undefined,
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* ----- Modes ----- */
  const isDisplayForm = (FormMode as number | undefined) === 4;
  const isNumber = type === 'number';
  const isFile = type === 'file';

  /* ----- Disabled/Hidden baseline (overridable) ----- */
  const baseDisabled = isDisplayForm;
  const baseHidden = false;

  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(baseDisabled);
  const [isDisabled, _setIsDisabled] = React.useState<boolean>(defaultDisable || !!submitting);
  const [isHidden, _setIsHidden] = React.useState<boolean>(baseHidden);

  // Guarded setters prevent no-op updates (which would re-render)
  const setDisabledIfChanged = React.useCallback((next: boolean) => {
    _setIsDisabled(prev => (prev !== next ? next : prev));
  }, []);
  const setDefaultDisableIfChanged = React.useCallback((next: boolean) => {
    setDefaultDisable(prev => (prev !== next ? next : prev));
  }, []);
  const setHiddenIfChanged = React.useCallback((next: boolean) => {
    _setIsHidden(prev => (prev !== next ? next : prev));
  }, []);

  // If FormMode flips (e.g., display mode), align the defaults (guarded)
  React.useEffect(() => {
    setDefaultDisableIfChanged(baseDisabled);
  }, [baseDisabled, setDefaultDisableIfChanged]);

  // Hidden baseline rarely changes; if it does, keep it guarded
  React.useEffect(() => {
    setHiddenIfChanged(baseHidden);
  }, [baseHidden, setHiddenIfChanged]);

  // Persist disabled across submit cycles if inherently disabled
  React.useEffect(() => {
    if (defaultDisable === false) setDisabledIfChanged(!!submitting);
    else setDisabledIfChanged(true);
  }, [defaultDisable, submitting, setDisabledIfChanged]);

  /* ----- Required flag ----- */
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  React.useEffect(() => setIsRequired(!!requiredProp), [requiredProp]);

  /* ----- Value, error, touched ----- */
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  /* ----- Expose DOM node via GlobalRefs ----- */
  const elemRef = React.useRef<HTMLInputElement>(null);

  // IMPORTANT: Call GlobalRefs ONCE on mount; unstable function refs can cause loops
  React.useEffect(() => {
    (GlobalRefs as ((el: HTMLElement | undefined) => void) | undefined)?.(elemRef.current ?? undefined);
    return () => (GlobalRefs as ((el: HTMLElement | undefined) => void) | undefined)?.(undefined);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // mount-only

  // Log only when flags actually change
  React.useEffect(() => {
    if (!DEBUG) return;
    // eslint-disable-next-line no-console
    console.log('[SingleLineComponent] state change', {
      id,
      displayName,
      isHidden,
      isDisabled,
      refPresent: !!elemRef.current,
    });
  }, [isHidden, isDisabled, displayName, id]);

  /* ----- Centralized rules: call once per field ID (guard updates inside) ----- */
  React.useEffect(() => {
    const formFieldProps: FormFieldsProps = {
      disabledList: AllDisabledFieldsNorm,
      hiddenList: AllHiddenFieldsNorm,
      userBasedList: userBasedPerms,
      curUserList: curUserInfo,
      curField: id,
      formStateData: FormData,
      listColumns: listCols,
    } as any;

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log(`[SingleLineComponent][${id}] formFieldProps`, formFieldProps);
    }

    let results: Array<{ isDisabled?: boolean; isHidden?: boolean }> = [];
    try {
      results = (formFieldsSetup(formFieldProps) as any) || [];
      if (!Array.isArray(results)) results = [];
    } catch (err) {
      // eslint-disable-next-line no-console
      console.warn(`[SingleLineComponent][${id}] formFieldsSetup threw`, err);
      results = [];
    }

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log(`[SingleLineComponent][${id}] formFieldsSetup results`, results);
    }

    if (results.length > 0) {
      for (let i = 0; i < results.length; i++) {
        if (results[i].isDisabled !== undefined) {
          const v = !!results[i].isDisabled;
          setDisabledIfChanged(v);
          setDefaultDisableIfChanged(v);
        }
        if (results[i].isHidden !== undefined) {
          setHiddenIfChanged(!!results[i].isHidden);
        }
      }
    }
    // Keep deps tight. If you truly need it reactive to permissions/data,
    // add specific deps back one by one (the guarded setters will still prevent loops).
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id]);

  /* ----- Number helpers ----- */
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  const decimalLimit: 1 | 2 | undefined = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return undefined; // automatic
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
    if (i !== -1) out = out.slice(0, i + 1) + out.slice(i + 1).replace(/\./g, '');
    return out;
  }, [allowNegative]);

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

  /* ----- Validation ----- */
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
    if (decimalLimit !== undefined && fractionDigits(val) > decimalLimit) return decimalLimitMsg(decimalLimit);
    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;
    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  const validate = React.useCallback((val: string): string => (
    isNumber ? validateNumber(val) : validateText(val)
  ), [isNumber, validateNumber, validateText]);

  /* ----- Prefill on mount (New vs Edit) ----- */
  React.useEffect(() => {
    const fromNew = (FormMode as number | undefined) === 8;
    const raw = fromNew
      ? (starterValue ?? '')
      : (FormData ? (FormData as Record<string, unknown>)[id] ?? '' : '');
    const str = raw === null || raw === undefined ? '' : String(raw);
    const sanitized0 = isNumber ? sanitizeDecimal(str) : str;
    const { value: sanitized } = isNumber ? applyDecimalLimit(sanitized0) : { value: sanitized0, trimmed: false as any };
    setLocalVal(sanitized);
    setError('');
    setTouched(false);

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log('[SingleLineComponent] prefill', { id, fromNew, raw, sanitized });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // once

  /* ----- Submit finalize (validate + commit if no blur yet) ----- */
  React.useEffect(() => {
    if (!submitting) return;
    const valueForValidation = isFile ? splitExt(localVal).base : localVal;
    const tooLong = !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? `Maximum length is ${maxLength} characters.` : validate(valueForValidation);

    setTouched(true);
    setError(finalError);
    (GlobalErrorHandle as (id: string, e: string | null) => void)?.(id, finalError === '' ? null : finalError);

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log('[SingleLineComponent] submit finalize', { id, localVal, valueForValidation, finalError });
    }

    if (isNumber) {
      const t = localVal.trim();
      (GlobalFormData as (id: string, v: unknown) => void)?.(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      (GlobalFormData as (id: string, v: unknown) => void)?.(id, out === '' ? null : out);
    }
  }, [submitting]); // when submit toggles

  /* ----- Selection helper (for paste logic) ----- */
  const getSelectionRange = (el: HTMLInputElement): { start: number; end: number } => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  /* ----- Paste handlers ----- */
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
      if (touched) setError(`Maximum length is ${maxLength} characters.`);
      return;
    }

    if (pasteText.length > spaceLeft) {
      e.preventDefault();
      const insert = pasteText.slice(0, Math.max(0, spaceLeft));
      const nextBase = input.value.slice(0, start) + insert + input.value.slice(end);

      if (isFile) {
        const { ext } = splitExt(localVal);
        const nextValue = nextBase === '' ? '' : `${nextBase}${ext}`;
        setLocalVal(nextValue);
      } else {
        setLocalVal(nextBase);
      }

      if (touched) setError(`Maximum length is ${maxLength} characters.`);
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
      if (touched) setError(decimalLimitMsg(decimalLimit));
    }
  };

  /* ----- Change/Blur ----- */
  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data): void => {
    const raw = data.value ?? '';

    if (isNumber) {
      const sanitized0 = sanitizeDecimal(raw);
      const { value: next, trimmed } = applyDecimalLimit(sanitized0);
      setLocalVal(next);
      if (touched) {
        if (trimmed && decimalLimit !== undefined) setError(decimalLimitMsg(decimalLimit));
        else setError(validateNumber(next));
      }
      return;
    }

    if (isFile) {
      const { ext } = splitExt(localVal);
      const recombined = raw === '' ? '' : `${raw}${ext}`;
      setLocalVal(recombined);
      if (touched) setError(validateText(recombined));
      return;
    }

    // TEXT
    setLocalVal(raw);
    if (touched) {
      if (isDefined(maxLength) && raw.length > maxLength) {
        setError(`Maximum length is ${maxLength} characters.`);
      } else {
        setError(validateText(raw));
      }
    }
  };

  const handleBlur: React.FocusEventHandler<HTMLInputElement> = (): void => {
    setTouched(true);

    const valueForValidation = isFile ? splitExt(localVal).base : localVal;
    const tooLong = !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? `Maximum length is ${maxLength} characters.` : validate(valueForValidation);

    setError(finalError);
    (GlobalErrorHandle as (id: string, e: string | null) => void)?.(id, finalError === '' ? null : finalError);

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log('[SingleLineComponent] blur commit', { id, localVal, valueForValidation, finalError });
    }

    if (isNumber) {
      const t = localVal.trim();
      (GlobalFormData as (id: string, v: unknown) => void)?.(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      (GlobalFormData as (id: string, v: unknown) => void)?.(id, out === '' ? null : out);
    }
  };

  /* ----- contentAfter & display value ----- */
  const extForAfter = isFile ? splitExt(localVal).ext : '';
  const after =
    (isFile && extForAfter)
      ? <Text size={400}>{extForAfter}</Text>
      : (isNumber && contentAfter === 'percentage')
        ? <Text size={400} id={`${id}Per`}>%</Text>
        : undefined;

  const displayValue = isFile ? splitExt(localVal).base : localVal;

  /* ----- Hidden? ----- */
  if (isHidden) {
    if (DEBUG) console.log(`[SingleLineComponent] rendering nothing because hidden: ${id}`);
    return <></>;
  }

  /* ----- Render ----- */
  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={error !== '' ? error : undefined}
      validationState={error !== '' ? 'error' : 'none'}
      hint={description}
    >
      <Input
        ref={elemRef}
        id={id}
        name={displayName}
        className={className}
        placeholder={placeholder}
        value={displayValue}
        onChange={handleChange}
        onBlur={handleBlur}
        onPaste={isNumber ? handleNumberPaste : handleTextPaste}
        disabled={isDisabled}
        maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}
        type={isNumber ? 'number' : 'text'}
        inputMode={isNumber ? 'decimal' : undefined}
        step="any"
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}
        contentAfter={after}
      />
    </Field>
  );
}