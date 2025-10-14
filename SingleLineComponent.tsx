/**
 * SingleLineComponent.tsx
 *
 * What this component does
 * ------------------------
 * Reusable single-line input built with Fluent UI v9 (<Field> + <Input>) that supports:
 *  • TEXT (default)
 *  • NUMBER (min/max, decimalPlaces, paste sanitization, optional “%” suffix via contentAfter)
 *  • FILE display (shows file base name in the field, and renders the extension in contentAfter)
 *
 * Integration points
 * ------------------
 * Works with a DynamicFormContext to:
 *  • Prefill from FormData / FormMode (New vs Edit)
 *  • Commit values to GlobalFormData on blur (and again during submit if focused)
 *  • Send validation messages to GlobalErrorHandle
 *  • Expose the input DOM node to GlobalRefs so other code can focus/scroll/etc.
 *  • Apply centralized form rules (disabled/hidden) via formFieldsSetup (matches ComboBox usage)
 *
 * Example usage
 * -------------
 * // TEXT (default)
 * <SingleLineComponent id="title" displayName="Title" maxLength={120} isRequired />
 *
 * // NUMBER
 * <SingleLineComponent
 *   id="discount"
 *   displayName="Discount"
 *   type="number"
 *   min={0}
 *   max={100}
 *   decimalPlaces="two"
 *   contentAfter="percentage"
 * />
 *
 * // FILE (display-only filename where extension is shown after the field)
 * <SingleLineComponent
 *   id="docName"
 *   displayName="Document Name"
 *   type="file"
 *   starterValue="Proposal_v3.docx"
 * />
 */

import * as React from 'react';
import { Field, Input, Text } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from '../Utilis/formFieldBased';

/* ───────────────────────────── Props ──────────────────────────── */
export interface SingleLineFieldProps {
  id: string;
  displayName: string;

  starterValue?: string | number;
  isRequired?: boolean;

  maxLength?: number;

  type?: 'text' | 'number' | 'file';

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

const decimalLimitMsg = (n: 1 | 2): string =>
  `Maximum ${n} decimal place${n === 1 ? '' : 's'} allowed.`;

const rangeMsg = (min?: number, max?: number): string =>
  (min !== undefined && max !== undefined)
    ? `Value must be between ${min} and ${max}.`
    : (min !== undefined)
      ? `Value must be ≥ ${min}.`
      : (max !== undefined)
        ? `Value must be ≤ ${max}.`
        : '';

const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

/** Split a file name into base and extension. */
function splitExt(name: string): { base: string; ext: string } {
  const i = name.lastIndexOf('.');
  if (i <= 0 || i === name.length - 1) return { base: name, ext: '' };
  return { base: name.slice(0, i), ext: name.slice(i) };
}

/* ────────── Types for strongly-typed context and rule results ────────── */

type DisabledHiddenList = unknown;

type GlobalFormDataFn = (id: string, v: unknown) => void;
type GlobalErrorHandleFn = (id: string, e: string | null) => void;
type GlobalRefsFn = (el: HTMLElement | undefined) => void;

interface ContextShape {
  FormData?: unknown;
  GlobalFormData?: GlobalFormDataFn;
  FormMode?: number;
  GlobalErrorHandle?: GlobalErrorHandleFn;
  GlobalRefs?: GlobalRefsFn;

  AllDisableFields?: DisabledHiddenList;
  AllDisabledFields?: DisabledHiddenList;

  AllHiddenFields?: DisabledHiddenList;

  userBasedPerms?: unknown;
  curUserInfo?: unknown;
  listCols?: unknown;
}

type RuleResult = { isDisabled?: boolean; isHidden?: boolean };

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

  /* Read the form context with a safe, typed shape (no `any`). */
  const ctx = React.useContext(DynamicFormContext) as unknown as ContextShape;

  const {
    FormData,
    GlobalFormData,
    FormMode,
    GlobalErrorHandle,
    GlobalRefs,

    AllDisableFields,
    AllDisabledFields,

    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
  } = ctx ?? {};

  const AllDisabledFieldsNorm: DisabledHiddenList | undefined =
    (AllDisableFields ?? AllDisabledFields) as DisabledHiddenList | undefined;
  const AllHiddenFieldsNorm: DisabledHiddenList | undefined =
    AllHiddenFields as DisabledHiddenList | undefined;

  /* Type flags */
  const isDisplayForm: boolean = (FormMode as number | undefined) === 4;
  const isNumber: boolean = type === 'number';
  const isFile: boolean = type === 'file';

  /* Disabled/hidden state with guarded setters to avoid no-op re-renders. */
  const baseDisabled: boolean = isDisplayForm;
  const baseHidden: boolean = false;

  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(baseDisabled);
  const [isDisabled, _setIsDisabled] = React.useState<boolean>(defaultDisable || !!submitting);
  const [isHidden, _setIsHidden] = React.useState<boolean>(baseHidden);

  const setDisabledIfChanged: (next: boolean) => void = React.useCallback((next: boolean): void => {
    _setIsDisabled(prev => (prev !== next ? next : prev));
  }, []);
  const setDefaultDisableIfChanged: (next: boolean) => void = React.useCallback((next: boolean): void => {
    setDefaultDisable(prev => (prev !== next ? next : prev));
  }, []);
  const setHiddenIfChanged: (next: boolean) => void = React.useCallback((next: boolean): void => {
    _setIsHidden(prev => (prev !== next ? next : prev));
  }, []);

  React.useEffect((): void => {
    setDefaultDisableIfChanged(baseDisabled);
  }, [baseDisabled, setDefaultDisableIfChanged]);

  React.useEffect((): void => {
    setHiddenIfChanged(baseHidden);
  }, [baseHidden, setHiddenIfChanged]);

  React.useEffect((): void => {
    if (defaultDisable === false) setDisabledIfChanged(!!submitting);
    else setDisabledIfChanged(true);
  }, [defaultDisable, submitting, setDisabledIfChanged]);

  /* Required flag mirrors prop. */
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  React.useEffect((): void => setIsRequired(!!requiredProp), [requiredProp]);

  /* Local value + validation state. */
  const [localVal, setLocalVal] = React.useState<string>(''); // stores the string the form will submit (full file name for FILE mode)
  const [error, setError] = React.useState<string>('');       // '' means no error
  const [touched, setTouched] = React.useState<boolean>(false);

  /* Expose the input DOM node to the outside via GlobalRefs (mount/unmount only). */
  const elemRef = React.useRef<HTMLInputElement>(null);
  React.useEffect((): (() => void) | void => {
    const fn = GlobalRefs as GlobalRefsFn | undefined;
    fn?.(elemRef.current ?? undefined);
    // cleanup
    return () => {
      fn?.(undefined);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* Centralized rules (disabled/hidden), aligned with ComboBox usage. */
  React.useEffect((): void => {
    // Normalize inputs to exactly what FormFieldsProps expects
    const disabledList = (AllDisabledFieldsNorm ?? {}) as Record<string, any>;
    const hiddenList = (AllHiddenFieldsNorm ?? {}) as Record<string, any>;
    const userBasedList = (userBasedPerms ?? {}) as Record<string, any>;
    const curUserList = (curUserInfo ?? {}) as Record<string, any>;
    const listColumns = Array.isArray(listCols) ? (listCols as string[]) : ([] as string[]);
    const formStateData = Array.isArray(FormData) ? (FormData as string[]) : ([] as string[]);

    const formFieldProps: FormFieldsProps = {
      disabledList,
      hiddenList,
      userBasedList,
      curUserList,
      curField: id,
      formStateData,
      listColumns,
    };

    let results: RuleResult[] = [];
    try {
      results = (formFieldsSetup(formFieldProps) as RuleResult[] | undefined) ?? [];
      if (!Array.isArray(results)) results = [];
    } catch {
      results = [];
    }

    if (results.length > 0) {
      for (let i = 0; i < results.length; i++) {
        const r = results[i];
        if (r.isDisabled !== undefined) {
          const v = !!r.isDisabled;
          setDisabledIfChanged(v);
          setDefaultDisableIfChanged(v);
        }
        if (r.isHidden !== undefined) {
          setHiddenIfChanged(!!r.isHidden);
        }
      }
    }
    // Keep deps tight to avoid loops from large, frequently-changing objects.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id]);

  /* Number helpers */
  const allowNegative: boolean = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  const decimalLimit: 1 | 2 | undefined = React.useMemo<1 | 2 | undefined>(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return undefined; // automatic
  }, [decimalPlaces]);

  const sanitizeDecimal: (s: string) => string = React.useCallback((s: string): string => {
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

  const applyDecimalLimit: (val: string) => { value: string; trimmed: boolean } = React.useCallback(
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

  /* Validation helpers */
  const validateText: (val: string) => string = React.useCallback((val: string): string => {
    if (isRequired && val.trim().length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired]);

  const isNumericString: (val: string) => boolean = React.useCallback((val: string): boolean => {
    if (!val || val.trim().length === 0) return false;
    const re = allowNegative ? /^-?(?:\d+\.?\d*|\.\d+)$/ : /^(?:\d+\.?\d*|\.\d+)$/;
    return re.test(val);
  }, [allowNegative]);

  const validateNumber: (val: string) => string = React.useCallback((val: string): string => {
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

  const validate: (val: string) => string = React.useCallback((val: string): string => (
    isNumber ? validateNumber(val) : validateText(val)
  ), [isNumber, validateNumber, validateText]);

  /* Prefill initial value on mount */
  React.useEffect((): void => {
    const fromNew = (FormMode as number | undefined) === 8;
    const rawUnknown = fromNew
      ? (starterValue ?? '')
      : (FormData && (FormData as Record<string, unknown>)[id] !== undefined
          ? (FormData as Record<string, unknown>)[id]
          : '');
    const str = rawUnknown === null || rawUnknown === undefined ? '' : String(rawUnknown);
    const sanitized0 = isNumber ? sanitizeDecimal(str) : str;

    const nextObj: { value: string; trimmed: boolean } = isNumber
      ? applyDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setLocalVal(nextObj.value);
    setError('');
    setTouched(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* Submit finalize */
  React.useEffect((): void => {
    if (!submitting) return;

    const valueForValidation = isFile ? splitExt(localVal).base : localVal;
    const tooLong = !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? `Maximum length is ${maxLength} characters.` : validate(valueForValidation);

    setTouched(true);
    setError(finalError);
    (GlobalErrorHandle as GlobalErrorHandleFn | undefined)?.(id, finalError === '' ? null : finalError);

    if (isNumber) {
      const t = localVal.trim();
      (GlobalFormData as GlobalFormDataFn | undefined)?.(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      (GlobalFormData as GlobalFormDataFn | undefined)?.(id, out === '' ? null : out);
    }
  }, [submitting, isNumber, isFile, localVal, maxLength, validate, GlobalErrorHandle, GlobalFormData, id]);

  /* Selection helper (used by paste logic) */
  const getSelectionRange = (el: HTMLInputElement): { start: number; end: number } => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  /* Paste handlers */
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

  /* Change/Blur */
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
    (GlobalErrorHandle as GlobalErrorHandleFn | undefined)?.(id, finalError === '' ? null : finalError);

    if (isNumber) {
      const t = localVal.trim();
      (GlobalFormData as GlobalFormDataFn | undefined)?.(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      (GlobalFormData as GlobalFormDataFn | undefined)?.(id, out === '' ? null : out);
    }
  };

  /* contentAfter & display value */
  const extForAfter: string = isFile ? splitExt(localVal).ext : '';
  const after: React.ReactNode =
    (isFile && extForAfter)
      ? <Text size={400}>{extForAfter}</Text>
      : (isNumber && contentAfter === 'percentage')
        ? <Text size={400} id={`${id}Per`}>%</Text>
        : undefined;

  const displayValue: string = isFile ? splitExt(localVal).base : localVal;

  /* Hidden fields render nothing. */
  if (isHidden) {
    return <></>;
  }

  /* Render */
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