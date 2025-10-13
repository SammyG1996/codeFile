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
 *  • Apply centralized rules via formFieldsSetup (hardened) with logs
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

/* ───────────────────────── Debug toggle ───────────────────────── */
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

/** Ensure a value is an object with `.items` array so `formFieldsSetup` can safely read `.items`. */
function asItems<T = unknown>(v: unknown): { items: T[] } {
  if (v && typeof v === 'object' && Array.isArray((v as any).items)) return v as { items: T[] };
  if (Array.isArray(v)) return { items: v as T[] };
  return { items: [] as T[] };
}

/** Lowercase & trim. */
const norm = (s: unknown): string => String(s ?? '').trim().toLowerCase();

/** SharePoint-ish encoding: replace spaces with `_x0020_` (basic heuristic). */
const toSPEncoded = (s: string): string => s.replace(/ /g, '_x0020_');

/** Remove all non-alphanumerics to compare loosely across variants. */
const toLoose = (s: string): string => s.replace(/[^a-z0-9]/gi, '').toLowerCase();

/** Does a list (strings / {items}) contain a name matching id/displayName/variants? */
function listHasName(list: unknown, id: string, displayName: string): boolean {
  const arr: string[] =
    Array.isArray(list) ? list.map(String)
    : (list && typeof list === 'object' && Array.isArray((list as any).items))
      ? (list as any).items.map(String)
      : [];
  const aliases = [
    id,
    displayName,
    toSPEncoded(displayName),
    toLoose(displayName),
    toLoose(id),
  ].map(norm);

  return arr.some(item => {
    const nItem = norm(item);
    return (
      aliases.includes(nItem) ||
      aliases.includes(norm(toLoose(item)))
    );
  });
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

  /* ----- Context (loosely typed on purpose) ----- */
  const formCtx = React.useContext(DynamicFormContext);
  const ctx = formCtx as unknown as Record<string, unknown>;

  const FormData = (ctx.FormData as Record<string, unknown> | undefined) ?? undefined;
  const FormMode = (ctx.FormMode as number | undefined) ?? undefined;
  const GlobalFormData = ctx.GlobalFormData as (id: string, value: unknown) => void;
  const GlobalErrorHandle = ctx.GlobalErrorHandle as (id: string, error: string | null) => void;
  const GlobalRefs = (ctx.GlobalRefs as ((el: HTMLElement | undefined) => void) | undefined) ?? undefined;

  const AllDisabledFields = ctx.AllDisabledFields; // could be array or {items:[]}
  const AllHiddenFields = ctx.AllHiddenFields;     // same
  const userBasedPerms = ctx.userBasedPerms;
  const curUserInfo = ctx.curUserInfo;
  const listCols = ctx.listCols;

  /* ----- Modes ----- */
  const isDisplayForm = FormMode === 4;
  const isNumber = type === 'number';
  const isFile = type === 'file';

  /* ----- Disabled/Hidden baseline (overridable) ----- */
  const baseDisabled = isDisplayForm; // original baseline rule
  const baseHidden = false;

  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(baseDisabled);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(defaultDisable || !!submitting);
  const [isHidden, setIsHidden] = React.useState<boolean>(baseHidden);

  React.useEffect(() => setDefaultDisable(baseDisabled), [baseDisabled]);
  React.useEffect(() => setIsHidden(baseHidden), [baseHidden]);

  // Persist disabled across submit cycles if inherently disabled
  React.useEffect(() => {
    if (defaultDisable === false) setIsDisabled(!!submitting);
    else setIsDisabled(true);
  }, [defaultDisable, submitting]);

  /* ----- Required flag ----- */
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  React.useEffect(() => setIsRequired(!!requiredProp), [requiredProp]);

  /* ----- Value, error, touched ----- */
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  /* ----- Expose DOM node via GlobalRefs ----- */
  const elemRef = React.useRef<HTMLInputElement>(null);
  React.useEffect(() => {
    GlobalRefs?.(elemRef.current ?? undefined);
    return () => GlobalRefs?.(undefined);
  }, []); // once

  // Log on mount (ref + displayName + initial hidden/disabled)
  React.useEffect(() => {
    if (!DEBUG) return;
    // eslint-disable-next-line no-console
    console.log('[SingleLineComponent] mount', {
      id,
      displayName,
      refPresent: !!elemRef.current,
      ref: elemRef.current,
      isHidden,
      isDisabled,
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // mount only

  // Log whenever isHidden / isDisabled changes
  React.useEffect(() => {
    if (!DEBUG) return;
    // eslint-disable-next-line no-console
    console.log('[SingleLineComponent] state change', {
      id,
      displayName,
      isHidden,
      isDisabled,
      refPresent: !!elemRef.current,
      ref: elemRef.current,
    });
  }, [isHidden, isDisabled, displayName, id]);

  /* ----- Centralized rules: formFieldsSetup (hardened + logs) ----- */
  React.useEffect(() => {
    const propsForSetup: FormFieldsProps = {
      disabledList: asItems(AllDisabledFields),
      hiddenList:   asItems(AllHiddenFields),
      userBasedList:asItems(userBasedPerms),
      curUserList:  asItems(curUserInfo),
      curField:     id,
      formStateData:asItems(FormData),  // some versions expect .items on this too
      listColumns:  asItems(listCols),
    } as any;

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log('[SingleLineComponent] setup input', {
        id, displayName,
        disabledList: propsForSetup.disabledList,
        hiddenList: propsForSetup.hiddenList,
        userBasedList: propsForSetup.userBasedList,
        curUserList: propsForSetup.curUserList,
        formStateData: propsForSetup.formStateData,
        listColumns: propsForSetup.listColumns,
      });
    }

    let appliedFromSetup = false;

    try {
      const raw = typeof formFieldsSetup === 'function' ? formFieldsSetup(propsForSetup) : [];
      const results = Array.isArray(raw) ? raw : [];
      if (DEBUG) {
        // eslint-disable-next-line no-console
        console.log('[SingleLineComponent] setup results', results);
      }

      for (const r of results) {
        if (!r || typeof r !== 'object') continue;
        if ('isDisabled' in r && r.isDisabled !== undefined) {
          const v = !!(r as any).isDisabled;
          appliedFromSetup = true;
          setIsDisabled(v);
          setDefaultDisable(v);
          if (DEBUG) console.log(`[SingleLineComponent] applied isDisabled=${v} from formFieldsSetup for ${id}`);
        }
        if ('isHidden' in r && r.isHidden !== undefined) {
          appliedFromSetup = true;
          setIsHidden(!!(r as any).isHidden);
          if (DEBUG) console.log(`[SingleLineComponent] applied isHidden=${(r as any).isHidden} from formFieldsSetup for ${id}`);
        }
      }
    } catch (err) {
      // eslint-disable-next-line no-console
      console.warn('formFieldsSetup failed for field', id, err);
    }

    // Fallback: if setup didn’t explicitly return rules, derive from static lists.
    if (!appliedFromSetup) {
      const derivedDisable = listHasName(AllDisabledFields, id, displayName);
      const derivedHidden  = listHasName(AllHiddenFields, id, displayName);

      if (DEBUG) {
        // eslint-disable-next-line no-console
        console.log('[SingleLineComponent] fallback list check', {
          id, displayName,
          derivedDisable, derivedHidden,
          disabledList: AllDisabledFields,
          hiddenList: AllHiddenFields,
        });
      }

      if (derivedDisable) {
        setIsDisabled(true);
        setDefaultDisable(true);
        if (DEBUG) console.log(`[SingleLineComponent] fallback applied isDisabled=true for ${id}`);
      }
      if (derivedHidden) {
        setIsHidden(true);
        if (DEBUG) console.log(`[SingleLineComponent] fallback applied isHidden=true for ${id}`);
      }
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [AllDisabledFields, AllHiddenFields, userBasedPerms, curUserInfo, id, FormData, listCols]);

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
    const fromNew = FormMode === 8;
    const raw = fromNew
      ? (starterValue ?? '')
      : (FormData ? (FormData[id] ?? '') : '');
    const str = raw === null || raw === undefined ? '' : String(raw);
    const sanitized0 = isNumber ? sanitizeDecimal(str) : str;
    const { value: sanitized } = isNumber ? applyDecimalLimit(sanitized0) : { value: sanitized0 };
    setLocalVal(sanitized);
    setError('');
    setTouched(false);

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log('[SingleLineComponent] prefill', { id, fromNew, raw, sanitized });
    }
  }, []); // once

  /* ----- Submit finalize (validate + commit if no blur yet) ----- */
  React.useEffect(() => {
    if (!submitting) return;
    const valueForValidation = isFile ? splitExt(localVal).base : localVal;
    const tooLong = !isNumber && isDefined(maxLength) && valueForValidation.length > (maxLength ?? Infinity);
    const finalError = tooLong ? `Maximum length is ${maxLength} characters.` : validate(valueForValidation);

    setTouched(true);
    setError(finalError);
    GlobalErrorHandle(id, finalError === '' ? null : finalError);

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log('[SingleLineComponent] submit finalize', { id, localVal, valueForValidation, finalError });
    }

    if (isNumber) {
      const t = localVal.trim();
      GlobalFormData(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      GlobalFormData(id, out === '' ? null : out);
    }
  }, [submitting]);

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
    GlobalErrorHandle(id, finalError === '' ? null : finalError);

    if (DEBUG) {
      // eslint-disable-next-line no-console
      console.log('[SingleLineComponent] blur commit', { id, localVal, valueForValidation, finalError });
    }

    if (isNumber) {
      const t = localVal.trim();
      GlobalFormData(id, t === '' ? null : (Number.isNaN(Number(t)) ? null : Number(t)));
    } else {
      const out = localVal.trim();
      GlobalFormData(id, out === '' ? null : out);
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
      hint={description}  /* v9: use `hint`, not `description` */
    >
      <Input
        ref={elemRef}
        id={id}                 /* per requirement: use props.id */
        name={displayName}      /* per requirement: use displayName */
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
        type={isNumber ? 'number' : 'text'}   // FILE renders as text (we only display the name)
        inputMode={isNumber ? 'decimal' : undefined}
        step="any"
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}

        contentAfter={after}
      />
    </Field>
  );
}