/**
 * SingleLineComponent.tsx
 *
 * What this component does
 * ------------------------
 * A reusable single-line Fluent UI input (<Field> + <Input>) that supports:
 *  • TEXT fields (default)
 *  • NUMBER fields (min/max, decimalPlaces, paste sanitization, optional % after the input)
 *  • FILE display fields (shows the file base name in the input and the extension in contentAfter)
 *
 * Integrates with DynamicFormContext to:
 *  • Read initial values (FormData / FormMode)
 *  • Commit values via GlobalFormData (on blur and when submitting)
 *  • Report validation via GlobalErrorHandle
 *  • Expose a ref via GlobalRefs
 *  • Respect centralized rules via formFieldsSetup (disabled/hidden), safely
 *
 * Example usage
 * -------------
 * // TEXT
 * <SingleLineComponent
 *   id="title"
 *   displayName="Title"
 *   starterValue="Initial value"
 *   isRequired
 *   maxLength={120}
 *   placeholder="Enter title"
 * />
 *
 * // NUMBER (shows % after the field)
 * <SingleLineComponent
 *   id="discount"
 *   displayName="Discount"
 *   type="number"
 *   min={0}
 *   max={100}
 *   decimalPlaces="two"
 *   contentAfter="percentage"
 *   starterValue={12.5}
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
import { formFieldsSetup, FormFieldsProps } from '../Utils/formFieldBased';

/* ============================================================================
   Props
   ========================================================================== */

export interface SingleLineFieldProps {
  id: string;
  displayName: string;

  starterValue?: string | number;
  isRequired?: boolean;

  // TEXT/FILE
  maxLength?: number;

  // Type selector (TEXT is default when omitted)
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

/* ============================================================================
   Helper constants & utilities
   ========================================================================== */

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

/** Split filename into base + extension (ext includes the dot, e.g., ".docx"). */
function splitExt(name: string): { base: string; ext: string } {
  const i = name.lastIndexOf('.');
  if (i <= 0 || i === name.length - 1) return { base: name, ext: '' };
  return { base: name.slice(0, i), ext: name.slice(i) };
}

/* ============================================================================
   Component
   ========================================================================== */

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

  /* ----- Context (read safely) ------------------------------------------------ */
  const formCtx = React.useContext(DynamicFormContext);
  const ctx = formCtx as unknown as Record<string, unknown>;

  const FormData = (ctx.FormData as Record<string, unknown> | undefined) ?? undefined;
  const FormMode = (ctx.FormMode as number | undefined) ?? undefined;

  const GlobalFormData = ctx.GlobalFormData as (id: string, value: unknown) => void;
  const GlobalErrorHandle = ctx.GlobalErrorHandle as (id: string, error: string | null) => void;
  const GlobalRefs = (ctx.GlobalRefs as ((el: HTMLElement | undefined) => void) | undefined) ?? undefined;

  const AllDisabledFields = ctx.AllDisabledFields;
  const AllHiddenFields = ctx.AllHiddenFields;
  const userBasedPerms = ctx.userBasedPerms;
  const curUserInfo = ctx.curUserInfo;
  const listCols = ctx.listCols;

  /* ----- Type mode flags ------------------------------------------------------ */
  const isDisplayForm = FormMode === 4;
  const isNumber = type === 'number';
  const isFile = type === 'file';

  /* ----- Disabled/Hidden (baseline) ------------------------------------------ */
  // Baseline rules — these can be overridden by centralized setup below
  const baseDisabled = isDisplayForm; // (your original baseline)
  const baseHidden = false;

  // Keep disable/hidden stateful so the setup can override them
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(baseDisabled);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(defaultDisable || !!submitting);
  const [isHidden, setIsHidden] = React.useState<boolean>(baseHidden);

  React.useEffect(() => setDefaultDisable(baseDisabled), [baseDisabled]);
  React.useEffect(() => setIsHidden(baseHidden), [baseHidden]);

  // Persist disabled through submit cycles: if inherently disabled, always true
  React.useEffect(() => {
    if (defaultDisable === false) setIsDisabled(!!submitting);
    else setIsDisabled(true);
  }, [defaultDisable, submitting]);

  /* ----- Required flag -------------------------------------------------------- */
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  React.useEffect(() => setIsRequired(!!requiredProp), [requiredProp]);

  /* ----- Local Value / Error state ------------------------------------------- */
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  /* ----- Expose ref to hosting app ------------------------------------------- */
  const elemRef = React.useRef<HTMLInputElement>(null);
  React.useEffect(() => {
    GlobalRefs?.(elemRef.current ?? undefined);
    return () => GlobalRefs?.(undefined);
  }, []); // run once

  /* ----- Centralized rules: formFieldsSetup (SAFE) ---------------------------- */
  React.useEffect(() => {
    // Normalize potentially missing shapes so helper won’t crash
    const safeListColumns =
      (listCols && typeof (listCols as any) === 'object' && Array.isArray((listCols as any).items))
        ? (listCols as any)
        : ({ items: [] } as any);

    const formFieldProps: FormFieldsProps = {
      disabledList: (AllDisabledFields ?? []) as unknown,
      hiddenList: (AllHiddenFields ?? []) as unknown,
      userBasedList: (userBasedPerms ?? []) as unknown,
      curUserList: (curUserInfo ?? []) as unknown,
      curField: id,
      formStateData: (FormData ?? {}) as any,
      listColumns: safeListColumns,
    } as any;

    try {
      const raw = typeof formFieldsSetup === 'function' ? formFieldsSetup(formFieldProps) : [];
      const results = Array.isArray(raw) ? raw : [];

      for (const r of results) {
        if (!r || typeof r !== 'object') continue;
        if ('isDisabled' in r && r.isDisabled !== undefined) {
          const v = !!(r as any).isDisabled;
          setIsDisabled(v);
          setDefaultDisable(v);
        }
        if ('isHidden' in r && r.isHidden !== undefined) {
          setIsHidden(!!(r as any).isHidden);
        }
      }
    } catch (err) {
      // Don’t let centralized setup crash the field.
      // eslint-disable-next-line no-console
      console.warn('formFieldsSetup failed for field', id, err);
    }
  }, [AllDisabledFields, AllHiddenFields, userBasedPerms, curUserInfo, id, FormData, listCols]);

  /* ----- Number helpers ------------------------------------------------------- */
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

  /* ----- Validation helpers --------------------------------------------------- */
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

  /* ----- Prefill on mount (New vs Edit) -------------------------------------- */
  React.useEffect(() => {
    const fromNew = FormMode === 8;
    const raw = fromNew
      ? (starterValue ?? '')
      : (FormData ? (FormData[id] ?? '') : '');

    const initial = raw === null || raw === undefined ? '' : String(raw);
    const sanitized0 = isNumber ? sanitizeDecimal(initial) : initial;
    const { value: sanitized } = isNumber ? applyDecimalLimit(sanitized0) : { value: sanitized0 };
    setLocalVal(sanitized);
    setError('');
    setTouched(false);
  }, []); // once

  /* ----- Submit finalize (validate+commit if no blur) ------------------------- */
  React.useEffect(() => {
    if (!submitting) return;
    const valueForValidation = isFile ? splitExt(localVal).base : localVal;
    // Max length check (TEXT/FILE)
    const tooLong =
      !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? `Maximum length is ${maxLength} characters.` : validate(valueForValidation);

    setTouched(true);
    setError(finalError);
    GlobalErrorHandle(id, finalError === '' ? null : finalError);

    // Commit the stored value (numbers as numbers; text/file as trimmed string or null)
    if (isNumber) {
      const t = localVal.trim();
      GlobalFormData(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      GlobalFormData(id, out === '' ? null : out);
    }
  }, [submitting]);

  /* ----- Misc input helpers --------------------------------------------------- */
  const getSelectionRange = (el: HTMLInputElement): { start: number; end: number } => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  /* ----- Paste handlers ------------------------------------------------------- */
  // TEXT/FILE: respect maxLength (FILE counts base name)
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

  // NUMBER: sanitize and enforce decimal places
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

  /* ----- onChange / onBlur ---------------------------------------------------- */
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
      // Input shows base only — keep localVal as base+ext
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
    const tooLong =
      !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? `Maximum length is ${maxLength} characters.` : validate(valueForValidation);

    setError(finalError);
    GlobalErrorHandle(id, finalError === '' ? null : finalError);

    // Commit once on blur
    if (isNumber) {
      const t = localVal.trim();
      GlobalFormData(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      GlobalFormData(id, out === '' ? null : out);
    }
  };

  /* ----- Build contentAfter + display value ---------------------------------- */
  const extForAfter = isFile ? splitExt(localVal).ext : '';
  const after =
    (isFile && extForAfter)
      ? <Text size={400}>{extForAfter}</Text>
      : (isNumber && contentAfter === 'percentage')
        ? <Text size={400} id={`${id}Per`}>%</Text>
        : undefined;

  const displayValue = isFile ? splitExt(localVal).base : localVal;

  /* ----- Hidden? -------------------------------------------------------------- */
  if (isHidden) return <></>;

  /* ----- Render --------------------------------------------------------------- */
  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={error !== '' ? error : undefined}
      validationState={error !== '' ? 'error' : 'none'}
    >
      <Input
        ref={elemRef}
        id={id}                 /* DOM id = props.id, per requirement */
        name={displayName}      /* input name uses the display label */
        className={className}
        placeholder={placeholder}
        value={displayValue}
        onChange={handleChange}
        onBlur={handleBlur}
        onPaste={isNumber ? handleNumberPaste : handleTextPaste}
        disabled={isDisabled}

        /* TEXT/FILE ONLY: let the browser stop extra typing; paste is handled above */
        maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}

        /* NUMBER ONLY */
        type={isNumber ? 'number' : 'text'}   // FILE renders as text; we only display a filename, not <input type="file">
        inputMode={isNumber ? 'decimal' : undefined}
        step="any"
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}

        contentAfter={after}
      />
      {description !== '' && <div className="descriptionText">{description}</div>}
    </Field>
  );
}