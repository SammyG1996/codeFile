/**
 * SingleLineComponent.tsx
 *
 * What this component is
 * ----------------------
 * A single-line input built with Fluent UI v9 (<Field> + <Input>) that supports:
 *   • TEXT (default)
 *   • NUMBER (type="number")
 *   • FILE display (type="file") → shows the file **base name** inside the input and the **extension** after the field
 *
 * How to use it
 * -------------
 * // TEXT (default)
 * <SingleLineComponent
 *   id="title"
 *   displayName="Title"
 *   starterValue="Initial value"
 *   isRequired
 *   maxLength={120}
 *   placeholder="Enter title"
 *   description="Short helper text"
 *   submitting={isSubmitting}
 * />
 *
 * // NUMBER (shows % after the field if contentAfter="percentage")
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
 *
 * // FILE display (editable base name; extension shown after the field)
 * <SingleLineComponent
 *   id="docName"
 *   displayName="Document Name"
 *   type="file"
 *   starterValue="Proposal_v3.docx"
 *   placeholder="Filename"
 * />
 */

import * as React from 'react';
import { Field, Input, Text } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

/* ════════════════════════
   Props
   ════════════════════════ */

export interface SingleLineFieldProps {
  /** Unique business identifier for this field (used for data/commits; NOT used directly as DOM id). */
  id: string;
  /** Human-friendly name for this field. Used for the label and the <Input name>. */
  displayName: string;

  /** Initial value for the field (text/number/file name). */
  starterValue?: string | number;
  /** Whether the field is required (validation happens on blur). */
  isRequired?: boolean;

  /** Max length for TEXT/FILE (count applies to the input text only; FILE counts base name). */
  maxLength?: number;

  /** Type selector. TEXT is default if omitted. */
  type?: 'text' | 'number' | 'file';

  /** Number-only constraints. */
  min?: number;
  max?: number;
  /** When set to 'percentage' and type==='number', a % is rendered after the field. */
  contentAfter?: 'percentage';
  /** Controls decimal precision for number typing/pasting. */
  decimalPlaces?: 'automatic' | 'one' | 'two';

  placeholder?: string;
  className?: string;
  /** Optional helper text under the input. */
  description?: string;

  /**
   * While submitting, controls are temporarily disabled.
   * If the field is inherently disabled by rules, that supersedes submitting (it stays disabled before/during/after submit).
   */
  submitting?: boolean;
}

/* ════════════════════════
   Helper constants & utils
   ════════════════════════ */

const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const INVALID_NUM_MSG = 'Please enter valid numeric value!';

const rangeMsg = (min?: number, max?: number): string =>
  (min !== undefined && max !== undefined)
    ? `Value must be between ${min} and ${max}.`
    : (min !== undefined)
      ? `Value must be ≥ ${min}.`
      : (max !== undefined)
        ? `Value must be ≤ ${max}.`
        : '';

const decimalLimitMsg = (n: 1 | 2): string =>
  `Maximum ${n} decimal place${n === 1 ? '' : 's'} allowed.`;

/** Narrowing helper: treat "defined" as "not undefined". */
const isDefined = <T,>(v: T | undefined): v is T => v !== undefined;

/** Safe reads from an unknown-shaped context object without using `any`. */
const hasKey = (obj: Record<string, unknown>, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(obj, key);
const getKey = <T,>(obj: Record<string, unknown>, key: string): T =>
  obj[key] as T;

/** Returns true if any of the provided keys on the object are truthy. */
const getCtxFlag = (obj: Record<string, unknown>, keys: string[]): boolean => {
  for (const k of keys) if (hasKey(obj, k)) return !!obj[k];
  return false;
};

/** Membership check against array / Set / comma string / object map. */
const toBool = (v: unknown): boolean => !!v;
const isListed = (bag: unknown, name: string): boolean => {
  const needle = name.trim().toLowerCase();
  if (!bag) return false;
  if (Array.isArray(bag)) return bag.some(v => String(v).trim().toLowerCase() === needle);
  if (typeof (bag as { has?: unknown }).has === 'function') {
    for (const v of (bag as Set<unknown>)) if (String(v).trim().toLowerCase() === needle) return true;
    return false;
  }
  if (typeof bag === 'string') return bag.split(',').map(s => s.trim().toLowerCase()).includes(needle);
  if (typeof bag === 'object') {
    for (const [k, v] of Object.entries(bag as Record<string, unknown>)) {
      if (k.trim().toLowerCase() === needle && toBool(v)) return true;
    }
    return false;
  }
  return false;
};

/**
 * Splits a filename into base and extension.
 * - Keeps the dot in the extension (e.g., ".docx")
 * - Returns empty extension for names like "README" and for leading-dot files without a trailing segment (".env" → base: ".env", ext: "")
 */
function splitExt(name: string): { base: string; ext: string } {
  const i = name.lastIndexOf('.');
  if (i <= 0 || i === name.length - 1) {
    return { base: name, ext: '' };
  }
  return { base: name.slice(0, i), ext: name.slice(i) };
}

/* ════════════════════════
   Component
   ════════════════════════ */

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

  // Pull needed pieces from context. We read keys safely rather than asserting a strict type.
  const formCtx = React.useContext(DynamicFormContext);
  const ctx = formCtx as unknown as Record<string, unknown>;

  const FormData = hasKey(ctx, 'FormData') ? getKey<Record<string, unknown>>(ctx, 'FormData') : undefined;
  const FormMode = hasKey(ctx, 'FormMode') ? getKey<number>(ctx, 'FormMode') : undefined;

  const GlobalFormData = getKey<(id: string, value: unknown) => void>(ctx, 'GlobalFormData');
  const GlobalErrorHandle = getKey<(id: string, error: string | null) => void>(ctx, 'GlobalErrorHandle');

  // Optional function to expose this input's DOM element to the hosting app.
  const GlobalRefs = hasKey(ctx, 'GlobalRefs')
    ? getKey<(el: HTMLElement | undefined) => void>(ctx, 'GlobalRefs')
    : undefined;

  // Type mode helpers
  const isDisplayForm = FormMode === 4;
  const isNumber = type === 'number';
  const isFile   = type === 'file';

  // Context-provided "disabled" & "visibility" rules
  const disabledFromCtx = getCtxFlag(ctx, ['isDisabled', 'disabled', 'formDisabled', 'Disabled']);
  const AllDisabledFields = hasKey(ctx, 'AllDisabledFields') ? ctx.AllDisabledFields : undefined;
  const AllHiddenFields   = hasKey(ctx, 'AllHiddenFields') ? ctx.AllHiddenFields : undefined;

  /**
   * Baseline (rule-based) state:
   * - baseDisabled: true if the field should be disabled by *rules* (form mode, context flags, lists)
   * - baseHidden:   true if the field should be hidden by *rules*
   */
  const baseDisabled = isDisplayForm || disabledFromCtx || isListed(AllDisabledFields, displayName);
  const baseHidden   = isListed(AllHiddenFields, displayName);

  /**
   * Remember whether this field is *inherently disabled* by rules.
   * This lets "submit-time disabling" *not* override real disabled rules.
   */
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(baseDisabled);
  React.useEffect(() => {
    setDefaultDisable(baseDisabled);
  }, [baseDisabled]);

  /**
   * Final `isDisabled` used by the Input.
   * If inherently disabled, it stays disabled before/during/after submit.
   * Otherwise, it disables only while submitting.
   */
  const [isDisabled, setIsDisabled] = React.useState<boolean>(defaultDisable || !!submitting);
  React.useEffect(() => {
    if (defaultDisable === false) {
      setIsDisabled(!!submitting);
    } else {
      setIsDisabled(true);
    }
  }, [defaultDisable, submitting]);

  /** Hidden state is derived from rules each render (no internal state needed). */
  const isHidden = baseHidden;

  /** Required flag can change from props. */
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  React.useEffect(() => setIsRequired(!!requiredProp), [requiredProp]);

  /** Local value stored as a string for rendering and validation (committed to GlobalFormData on blur). */
  const [localVal, setLocalVal] = React.useState<string>('');
  /** Error text to display under the field (set on blur and as-you-type where it makes sense). */
  const [error, setError] = React.useState<string>('');
  /** We only push errors to GlobalErrorHandle after the first blur. */
  const [touched, setTouched] = React.useState<boolean>(false);

  /** DOM ref to the actual input element for external access via GlobalRefs. */
  const elemRef = React.useRef<HTMLInputElement>(null);

  /**
   * Call GlobalRefs once on mount (and clean up on unmount).
   * IMPORTANT: Do not depend on GlobalRefs in the deps array; some providers recreate
   * that function every render, which can cause a loop if we depend on it.
   */
  React.useEffect(() => {
    GlobalRefs?.(elemRef.current ?? undefined);
    return () => {
      GlobalRefs?.(undefined);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // run once

  /* ───────────────────────── number helpers ───────────────────────── */

  /** Coerce unknown/nullable values to a safe string for the input. */
  const valueToString = (v: unknown): string =>
    (v === null || v === undefined) ? '' : String(v);

  // Determine if negative numbers are allowed by the range constraints
  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  // Decimal precision control for typing/pasting
  const decimalLimit: 1 | 2 | undefined = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return undefined; // automatic (no enforced limit)
  }, [decimalPlaces]);

  /** Remove illegal characters, normalize multiple '-' and '.', and honor negativity rules. */
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

  const lengthMsg = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

  const fractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
  };

  /**
   * Enforce at most N decimal places (if configured).
   * Returns the possibly-trimmed value and whether trimming happened.
   */
  const applyDecimalLimit = React.useCallback(
    (val: string): { value: string; trimmed: boolean } => {
      if (decimalLimit === undefined) return { value: val, trimmed: false };
      const dot = val.indexOf('.');
      if (dot === -1) return { value: val, trimmed: false };
      const whole = val.slice(0, dot + 1);
      const frac = val.slice(0, dot + 1) ? val.slice(dot + 1) : '';
      if (frac.length <= decimalLimit) return { value: val, trimmed: false };
      return { value: whole + frac.slice(0, decimalLimit), trimmed: true };
    },
    [decimalLimit]
  );

  /* ───────────────────────── validation ───────────────────────── */

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

    if (decimalLimit !== undefined && fractionDigits(val) > decimalLimit) {
      return decimalLimitMsg(decimalLimit);
    }

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  /** Dispatch to the correct validator based on type mode. */
  const validate = React.useCallback((val: string): string => (
    isNumber ? validateNumber(val) : validateText(val)
  ), [isNumber, validateNumber, validateText]);

  /* ───────────────────────── lifecycle: initial value ───────────────────────── */

  /**
   * On mount we prefill the local value from:
   *   - New Form (FormMode === 8): props.starterValue
   *   - Edit Form: context.FormData[id]
   * We do NOT push anything to GlobalFormData here—that only happens on blur.
   */
  React.useEffect((): void => {
    const fromNewMode = FormMode === 8;
    const raw = fromNewMode
      ? (starterValue !== undefined ? starterValue : '')
      : (FormData ? (FormData as Record<string, unknown>)[id] : '');

    const initial = valueToString(raw);
    const sanitized0 = isNumber ? sanitizeDecimal(initial) : initial;
    const { value: sanitized, trimmed } = isNumber
      ? applyDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setLocalVal(sanitized);
    setTouched(false);
    if (trimmed && decimalLimit !== undefined) {
      // Surface a warning if we had to trim decimals on load.
      setError(decimalLimitMsg(decimalLimit));
    } else {
      setError('');
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // run once

  /* ───────────────────────── handlers ───────────────────────── */

  /** Utility to read current selection; used by paste handlers. */
  const getSelectionRange = (el: HTMLInputElement): { start: number; end: number } => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  /** TEXT/FILE paste: enforce maxLength (TEXT counts raw, FILE counts base name). */
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
      if (touched) setError(lengthMsg);
      return;
    }

    if (pasteText.length > spaceLeft) {
      e.preventDefault();
      const insert = pasteText.slice(0, Math.max(0, spaceLeft));
      const nextBase = input.value.slice(0, start) + insert + input.value.slice(end);

      if (isFile) {
        // In file mode, we store base+ext in localVal, but the input displays only the base.
        const { ext } = splitExt(localVal);
        const nextValue = nextBase === '' ? '' : `${nextBase}${ext}`;
        setLocalVal(nextValue);
      } else {
        setLocalVal(nextBase);
      }

      if (touched) setError(lengthMsg);
    }
  };

  /** NUMBER paste: sanitize and optionally trim fractional digits. */
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

  /** onChange keeps local value only; we validate lightly while typing. */
  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data): void => {
    const raw = data.value ?? '';

    if (isNumber) {
      const sanitized0 = sanitizeDecimal(raw);
      const { value: next, trimmed } = applyDecimalLimit(sanitized0);
      setLocalVal(next);
      if (trimmed && decimalLimit !== undefined) {
        if (touched) setError(decimalLimitMsg(decimalLimit));
        return;
      }
      if (touched) setError(validateNumber(next));
      return;
    }

    if (isFile) {
      // User edits only the base name; we preserve the extension in storage.
      const { ext } = splitExt(localVal);
      const recombined = raw === '' ? '' : `${raw}${ext}`;
      setLocalVal(recombined);
      if (touched) setError(validateText(recombined));
      return;
    }

    // TEXT
    setLocalVal(raw);
    // IMPORTANT: equal to maxLength is allowed; only error if strictly greater.
    if (isDefined(maxLength) && raw.length > maxLength) {
      if (touched) setError(lengthMsg);
      return;
    }
    if (touched) setError(validateText(raw));
  };

  /**
   * onBlur:
   *  - mark as touched (so future changes surface validation inline)
   *  - compute final error (including maxLength)
   *  - push error to GlobalErrorHandle
   *  - commit value to GlobalFormData (null for empty)
   */
  const handleBlur: React.FocusEventHandler<HTMLInputElement> = (): void => {
    setTouched(true);

    const valueForValidation = isFile ? splitExt(localVal).base : localVal;

    // IMPORTANT: equal to maxLength is valid; only error if strictly greater.
    const tooLong = !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? lengthMsg : validate(valueForValidation);

    setError(finalError);
    GlobalErrorHandle(id, finalError === '' ? null : finalError);

    // Commit to form data: numbers as numbers, text/file as trimmed string or null
    if (isNumber) {
      const t = localVal.trim();
      GlobalFormData(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      GlobalFormData(id, out === '' ? null : out);
    }
  };

  /* ───────────────────────── render ───────────────────────── */

  // Use a DOM-safe, per-instance id to avoid collisions in the page
  const inputDomId = `${id}__input`;

  // contentAfter: show file extension if in FILE mode; otherwise support % for numbers
  const extForAfter = isFile ? splitExt(localVal).ext : '';
  const after = (isFile && extForAfter)
    ? <Text size={400} id={`${inputDomId}Ext`}>{extForAfter}</Text>
    : (isNumber && contentAfter === 'percentage')
      ? <Text size={400} id={`${inputDomId}Per`}>%</Text>
      : undefined;

  // What appears in the input: file base for FILE mode; raw local value otherwise
  const displayValue = isFile ? splitExt(localVal).base : localVal;

  return (
    <div hidden={isHidden} className="fieldClass">
      <Field
        label={displayName}
        required={isRequired}
        validationMessage={error !== '' ? error : undefined}
        validationState={error !== '' ? 'error' : undefined}
      >
        <Input
          ref={elemRef}
          id={inputDomId}     /* DOM id (namespaced to avoid collisions) */
          name={displayName}  /* business name (as you requested) */
          className={className}
          placeholder={placeholder}
          value={displayValue}
          onChange={handleChange}
          onBlur={handleBlur}
          onPaste={isNumber ? handleNumberPaste : handleTextPaste}
          disabled={isDisabled}

          /* TEXT/FILE ONLY: enforce length in the browser (TEXT counts raw, FILE counts base) */
          maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}

          /* NUMBER ONLY: render numeric input and provide hints for mobile keyboards */
          type={isNumber ? 'number' : 'text'}  // FILE renders as 'text' because we only display the name, not an actual <input type="file">
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