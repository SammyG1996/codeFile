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
import { formFieldsSetup, FormFieldsProps } from '../Utils/formFieldBased';

/* ───────────────────────────── Props ──────────────────────────── */
/** Public props accepted by the component. */
export interface SingleLineFieldProps {
  /** Unique identifier used for committing to GlobalFormData and for the input id attribute. */
  id: string;
  /** Human-friendly label shown above the field and used as the input name. */
  displayName: string;

  /** Initial value for new items (FormMode === 8). */
  starterValue?: string | number;
  /** Whether the field is required (simple non-empty check for text/file; number must parse). */
  isRequired?: boolean;

  /** Maximum length for TEXT/FILE input (enforced on paste and validated on blur/submit). */
  maxLength?: number;

  /** Field rendering type. Defaults to 'text' if omitted. */
  type?: 'text' | 'number' | 'file';

  /** Number-only validation range. */
  min?: number;
  max?: number;

  /** Number-only decimal places policy. */
  decimalPlaces?: 'automatic' | 'one' | 'two';

  /** Number-only contentAfter helper. When 'percentage', a trailing % is rendered. */
  contentAfter?: 'percentage';

  /** Standard appearance props. */
  placeholder?: string;
  className?: string;
  description?: string;

  /**
   * Form submit flag from parent.
   * When true, the component validates and commits current value to GlobalFormData.
   * Also used to disable the input unless the field is inherently disabled by rules.
   */
  submitting?: boolean;
}

/* ───────────────────── Helpers & messages ─────────────────────── */

/** User-facing validation messages. */
const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const INVALID_NUM_MSG = 'Please enter valid numeric value!';

/** Pretty messages for decimal and range limits. */
const decimalLimitMsg = (n: 1 | 2) => `Maximum ${n} decimal place${n === 1 ? '' : 's'} allowed.`;
const rangeMsg = (min?: number, max?: number) =>
  (min !== undefined && max !== undefined)
    ? `Value must be between ${min} and ${max}.`
    : (min !== undefined)
      ? `Value must be ≥ ${min}.`
      : (max !== undefined)
        ? `Value must be ≤ ${max}.`
        : '';

/** Type guard: treat “defined” as anything not strictly undefined. */
const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

/**
 * Split a file name into base and extension.
 * The returned `ext` retains the dot (e.g., ".docx"). If there is no extension,
 * or the dot is the first/last character (".env", "name."), `ext` is '' and `base` is the original.
 */
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

  /* ── Read the form context. We keep this permissive to avoid tight coupling to provider shape. ── */
  const ctx = React.useContext(DynamicFormContext) as Record<string, unknown>;

  // Common context values used throughout the app.
  const {
    FormData,                 // existing values when editing
    GlobalFormData,           // callback to write a value: (id, value) => void
    FormMode,                 // numeric mode (8 = New, 4 = Display, others = Edit)
    GlobalErrorHandle,        // callback to write an error: (id, message|null) => void
    GlobalRefs,               // callback to expose a ref: (element|undefined) => void

    // Note: some places provide "AllDisableFields" (without the 'd'), others "AllDisabledFields".
    AllDisableFields,
    AllDisabledFields,

    AllHiddenFields,          // list of fields hidden by external rules
    userBasedPerms,           // optional user permissions bag used by formFieldsSetup
    curUserInfo,              // current user info consumed by formFieldsSetup
    listCols,                 // list column metadata consumed by formFieldsSetup
  } = (ctx as any) ?? {};

  // Normalize the disabled list key so either spelling works.
  const AllDisabledFieldsNorm = (AllDisableFields ?? AllDisabledFields) as unknown;
  const AllHiddenFieldsNorm = AllHiddenFields as unknown;

  /* ── Type flags ── */
  const isDisplayForm = (FormMode as number | undefined) === 4;
  const isNumber = type === 'number';
  const isFile = type === 'file';

  /* ── Disabled/hidden state ─────────────────────────────────────────────────────────────
   * Two sources can affect disabled/hidden:
   *  1) A baseline from mode (e.g., display mode disables the field).
   *  2) Centralized rules returned by formFieldsSetup (and its supporting lists).
   *
   * We persist "inherent" disablement (defaultDisable) so that while submitting we keep
   * the field disabled if rules say it should be disabled even after submit finishes.
   * Guarded setters are used to avoid no-op state updates that would trigger re-renders.
   * ───────────────────────────────────────────────────────────────────────────────────── */
  const baseDisabled = isDisplayForm;
  const baseHidden = false;

  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(baseDisabled);
  const [isDisabled, _setIsDisabled] = React.useState<boolean>(defaultDisable || !!submitting);
  const [isHidden, _setIsHidden] = React.useState<boolean>(baseHidden);

  /** Guarded setters only update state when the next value is different. */
  const setDisabledIfChanged = React.useCallback((next: boolean) => {
    _setIsDisabled(prev => (prev !== next ? next : prev));
  }, []);
  const setDefaultDisableIfChanged = React.useCallback((next: boolean) => {
    setDefaultDisable(prev => (prev !== next ? next : prev));
  }, []);
  const setHiddenIfChanged = React.useCallback((next: boolean) => {
    _setIsHidden(prev => (prev !== next ? next : prev));
  }, []);

  /** Align defaults when display/edit mode changes. */
  React.useEffect(() => {
    setDefaultDisableIfChanged(baseDisabled);
  }, [baseDisabled, setDefaultDisableIfChanged]);

  /** Hidden baseline is constant; if it changes, update guarded. */
  React.useEffect(() => {
    setHiddenIfChanged(baseHidden);
  }, [baseHidden, setHiddenIfChanged]);

  /**
   * During submit, all fields are disabled by the parent form. When submit ends,
   * if this field is inherently disabled (by mode or centralized rules), it should remain disabled.
   */
  React.useEffect(() => {
    if (defaultDisable === false) setDisabledIfChanged(!!submitting);
    else setDisabledIfChanged(true);
  }, [defaultDisable, submitting, setDisabledIfChanged]);

  /* ── Required flag mirrors prop (could also be set by rules separately if desired). ── */
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  React.useEffect(() => setIsRequired(!!requiredProp), [requiredProp]);

  /* ── Local value + validation state ── */
  const [localVal, setLocalVal] = React.useState<string>(''); // actual committed string (full name for file)
  const [error, setError] = React.useState<string>('');       // current validation message ('' = none)
  const [touched, setTouched] = React.useState<boolean>(false);

  /* ── Expose the input DOM node to the outside via GlobalRefs (mount/unmount only). ── */
  const elemRef = React.useRef<HTMLInputElement>(null);
  React.useEffect(() => {
    (GlobalRefs as ((el: HTMLElement | undefined) => void) | undefined)?.(elemRef.current ?? undefined);
    return () => (GlobalRefs as ((el: HTMLElement | undefined) => void) | undefined)?.(undefined);
    // Intentionally mount-only: changing GlobalRefs reference shouldn’t cause loops here.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* ── Centralized rules (disabled/hidden), aligned with ComboBox usage ──
   * We build FormFieldsProps and call formFieldsSetup. If results are returned,
   * we apply `isDisabled` / `isHidden`. Setters are guarded to prevent loops.
   * Dependencies are intentionally kept small; add more if those inputs truly
   * change at runtime for your form.
   */
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

    let results: Array<{ isDisabled?: boolean; isHidden?: boolean }> = [];
    try {
      results = (formFieldsSetup(formFieldProps) as any) || [];
      if (!Array.isArray(results)) results = [];
    } catch {
      results = [];
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
    // Keep deps tight to avoid loops from large, frequently-changing objects.
    // Add specific deps back if rules must respond to them in real time.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id]);

  /* ── Number helpers ── */
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  /** Convert "automatic" | "one" | "two" into an optional 1 or 2 limit for decimals. */
  const decimalLimit: 1 | 2 | undefined = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return undefined; // automatic
  }, [decimalPlaces]);

  /** Keep only digits, at most one leading '-', and at most one '.'. Respect allowNegative. */
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

  /** Helper to count digits after the decimal. */
  const fractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
  };

  /** Trims the fractional part to the configured decimal limit, if set. */
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

  /* ── Validation helpers ── */
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

  /** Choose the correct validator based on field type. */
  const validate = React.useCallback((val: string): string => (
    isNumber ? validateNumber(val) : validateText(val)
  ), [isNumber, validateNumber, validateText]);

  /* ── Prefill initial value on mount ──
   * For New (FormMode === 8), use starterValue; otherwise read existing value from FormData.
   * Number values are sanitized/trimmed for decimal rules, text/file pass through.
   */
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
    // Mount-only: we don't want to fight user input by re-prefilling later.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* ── Submit finalize ──
   * When the parent toggles `submitting` to true, validate current value and commit it.
   * This ensures the latest value is captured even if the control didn't blur.
   */
  React.useEffect(() => {
    if (!submitting) return;

    // In FILE mode we validate the base (the user-editable part), not the base+ext pair.
    const valueForValidation = isFile ? splitExt(localVal).base : localVal;

    const tooLong = !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? `Maximum length is ${maxLength} characters.` : validate(valueForValidation);

    setTouched(true);
    setError(finalError);
    (GlobalErrorHandle as (id: string, e: string | null) => void)?.(id, finalError === '' ? null : finalError);

    // Commit to GlobalFormData with the correct type:
    //  • NUMBER: null for empty/invalid, otherwise Number
    //  • TEXT/FILE: null for empty string, otherwise the string (FILE stores full name base+ext)
    if (isNumber) {
      const t = localVal.trim();
      (GlobalFormData as (id: string, v: unknown) => void)?.(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      (GlobalFormData as (id: string, v: unknown) => void)?.(id, out === '' ? null : out);
    }
  }, [submitting, isNumber, isFile, localVal, maxLength, validate, GlobalErrorHandle, GlobalFormData, id]);

  /* ── Selection helper (used by paste logic) ── */
  const getSelectionRange = (el: HTMLInputElement): { start: number; end: number } => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  /* ── Paste handlers ──
   * TEXT/FILE: respect maxLength by trimming pasted text when necessary.
   * NUMBER: enforce sanitize + decimal limit during paste.
   */
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
        // For FILE mode, localVal stores the full name; recombine base + existing ext.
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

  /* ── Change/Blur ──
   * onChange updates local state and (optionally) shows validation while typing after the first blur.
   * onBlur validates and pushes both error and value to the global form handlers.
   */
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

    // Validate the user-editable portion (base) for FILE mode.
    const valueForValidation = isFile ? splitExt(localVal).base : localVal;
    const tooLong = !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? `Maximum length is ${maxLength} characters.` : validate(valueForValidation);

    setError(finalError);
    (GlobalErrorHandle as (id: string, e: string | null) => void)?.(id, finalError === '' ? null : finalError);

    // Commit to global form data after blur (single source of truth for writes while editing).
    if (isNumber) {
      const t = localVal.trim();
      (GlobalFormData as (id: string, v: unknown) => void)?.(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      (GlobalFormData as (id: string, v: unknown) => void)?.(id, out === '' ? null : out);
    }
  };

  /* ── contentAfter & display value ──
   * FILE: show extension outside the input; inside we display only the base name.
   * NUMBER: show trailing % when contentAfter === 'percentage'.
   */
  const extForAfter = isFile ? splitExt(localVal).ext : '';
  const after =
    (isFile && extForAfter)
      ? <Text size={400}>{extForAfter}</Text>
      : (isNumber && contentAfter === 'percentage')
        ? <Text size={400} id={`${id}Per`}>%</Text>
        : undefined;

  const displayValue = isFile ? splitExt(localVal).base : localVal;

  /* ── Hidden fields render nothing. This keeps the DOM clean and avoids tab stops. ── */
  if (isHidden) {
    return <></>;
  }

  /* ── Render ── */
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
        id={id}                 /* use the id prop so labels/automation can target this field */
        name={displayName}      /* set a readable name for form serializers and accessibility tooling */
        className={className}
        placeholder={placeholder}
        value={displayValue}
        onChange={handleChange}
        onBlur={handleBlur}
        onPaste={isNumber ? handleNumberPaste : handleTextPaste}
        disabled={isDisabled}

        /* TEXT/FILE ONLY */
        maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}

        /* NUMBER ONLY */
        type={isNumber ? 'number' : 'text'}   /* FILE renders as 'text' — only the filename base is editable */
        inputMode={isNumber ? 'decimal' : undefined}
        step="any"
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}

        /* Suffix content: file extension or % sign */
        contentAfter={after}
      />
    </Field>
  );
}