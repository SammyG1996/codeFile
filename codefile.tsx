/**
 * SingleLineComponent.tsx
 *
 * USAGE
 * -----
 * TEXT mode (default):
 *   <SingleLineComponent
 *     id="title"                     // REQUIRED
 *     displayName="Title"            // REQUIRED
 *     maxLength={120}                // OPTIONAL
 *     isRequired={true}              // OPTIONAL
 *     disabled={false}               // OPTIONAL (overridden by AllDisabledFields / perms)
 *     starterValue="Prefilled text"  // OPTIONAL (used in New mode)
 *     placeholder="Enter title"      // OPTIONAL
 *     description="Shown under input as helper text" // OPTIONAL
 *     className="w-full"             // OPTIONAL
 *   />
 *
 * NUMBER mode:
 *   <SingleLineComponent
 *     id="discount"                  // REQUIRED
 *     displayName="Discount"         // REQUIRED
 *     type="number"                  // REQUIRED for number mode
 *     min={0}                        // OPTIONAL (inclusive)
 *     max={100}                      // OPTIONAL (inclusive)
 *     decimalPlaces="two"            // OPTIONAL: 'automatic' | 'one' | 'two' (default 'automatic')
 *     contentAfter="percentage"      // OPTIONAL: renders '%' suffix
 *     isRequired={true}              // OPTIONAL
 *     disabled={false}               // OPTIONAL (overridden by AllDisabledFields / perms)
 *     starterValue={12.5}            // OPTIONAL (used in New mode)
 *     placeholder="e.g. 12.5"        // OPTIONAL
 *     description="0 - 100, up to 2 decimals" // OPTIONAL
 *     className="w-48"               // OPTIONAL
 *   />
 *
 * NOTES
 * - Prefill runs ONCE on mount.
 * - Permissions recompute when context/props change.
 * - Hiding logic is scaffolded but commented out for later enablement.
 */

import * as React from 'react';
import { Field, Input, Text, useId } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

/* -------------------------------------------------------
 * Local TS types for context (use until your context exports them)
 * -----------------------------------------------------*/
type GroupRule = {
  groupSource: {
    groupName: string;
    editableOnly?: string[];
    visibleOnly?: string[];
  };
};

interface UserBasedPerms {
  groupBased?: GroupRule[];
}

interface SPGroupEntry {
  title?: string;
  Title?: string;
  [k: string]: any;
}

interface CurUserInfo {
  displayName?: string;
  Dept?: string;
  SPGroups?: SPGroupEntry[];
  [k: string]: any;
}

interface DynamicFormContextType {
  FormData?: Record<string, any>;
  FormMode?: number;
  GlobalFormData: (elmIntrnName: string, elmValue: any) => void;
  GlobalErrorHandle: (elmIntrnName: string, errorMessage: string | null) => void;

  // NEW pieces we’re using
  AllDisabledFields?: string[];
  AllHiddenFields?: string[];
  userBasedPerms?: UserBasedPerms;
  curUserInfo?: CurUserInfo;
}

/* ----------------------------------------------------- */

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
  decimalPlaces?: 'automatic' | 'one' | 'two';

  placeholder?: string;
  className?: string;
  description?: string;
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
const norm = (s: unknown): string => String(s ?? '').trim().toLowerCase();

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

  // Cast the context to our local interface so TS knows about the added fields
  const {
    FormData,
    GlobalFormData,
    FormMode,
    GlobalErrorHandle,
    AllDisabledFields,
    AllHiddenFields, // used later (commented scaffold)
    userBasedPerms,
    curUserInfo,
  } = React.useContext(DynamicFormContext) as unknown as DynamicFormContextType;

  const inputId = useId('input');

  // Controlled state
  const [localVal, setLocalVal] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Mirror flags (reactive to prop changes)
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  // const [isHidden, setIsHidden] = React.useState<boolean>(false); // for future hiding

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  const isNumber = type === 'number';
  const toStr = (v: unknown) => (v === null || v === undefined ? '' : String(v));

  const allowNegative = (isDefined(min) && min < 0) || (isDefined(max) && max < 0);

  const decimalLimit: 1 | 2 | null = React.useMemo(() => {
    if (decimalPlaces === 'one') return 1;
    if (decimalPlaces === 'two') return 2;
    return null;
  }, [decimalPlaces]);

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

  // const digitsOnly = (s: string) => s.replace(/[^\d]/g, ''); // UNUSED (kept for reference)

  const lengthMsg = isDefined(maxLength) ? `Maximum length is ${maxLength} characters.` : '';

  const getFractionDigits = (val: string): number => {
    const dot = val.indexOf('.');
    return dot === -1 ? 0 : Math.max(0, val.length - dot - 1);
  };

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

    if (decimalLimit !== null && getFractionDigits(val) > decimalLimit) {
      return decimalLimitMsg(decimalLimit);
    }

    const n = Number(val);
    if (Number.isNaN(n)) return INVALID_NUM_MSG;

    if (isDefined(min) && n < min) return rangeMsg(min, max);
    if (isDefined(max) && n > max) return rangeMsg(min, max);
    return '';
  }, [isRequired, min, max, isNumericString, decimalLimit]);

  const computeError = React.useCallback(
    (val: string) => (isNumber ? validateNumber(val) : validateText(val)),
    [isNumber, validateNumber, validateText]
  );

  const commitValue = React.useCallback((val: string, err: string) => {
    GlobalErrorHandle(id, err ? err : null);
    GlobalFormData(id, val);
  }, [GlobalErrorHandle, GlobalFormData, id]);

  // Prefill ONCE on mount
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
      GlobalFormData(id, sanitized);
    } else {
      const existing = FormData ? toStr((FormData as any)[id]) : '';
      const sanitized0 = isNumber ? decimalSanitizer(existing) : existing;
      const { value: sanitized, trimmed } = isNumber
        ? enforceDecimalLimit(sanitized0)
        : { value: sanitized0, trimmed: false };
      setLocalVal(sanitized);
      setError(trimmed && decimalLimit !== null ? decimalLimitMsg(decimalLimit) : '');
      setTouched(false);
      GlobalFormData(id, sanitized);
    }
    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // run once

  /**
   * Permissions / visibility effect
   * Precedence:
   *   1) AllDisabledFields → disabled (hard stop)
   *   2) userBasedPerms + curUserInfo groups → decide disabled (and future hidden)
   *   3) Prop-level disabled flag still applies if true
   */
  React.useEffect(() => {
    const fieldNameNorm = norm(displayName);

    // 1) AllDisabledFields precedence
    const disabledByAllDisabled =
      Array.isArray(AllDisabledFields) &&
      AllDisabledFields.some((n) => norm(n) === fieldNameNorm);

    if (disabledByAllDisabled) {
      setIsDisabled(true);
      return;
    }

    // 2) Group-based perms
    const groupsArray: SPGroupEntry[] = Array.isArray(curUserInfo?.SPGroups)
      ? (curUserInfo!.SPGroups as SPGroupEntry[])
      : [];
    const userGroupTitles = new Set(
      groupsArray
        .map((g) => norm(g?.title ?? g?.Title))
        .filter((s) => s !== '')
    );

    const rules: GroupRule[] = Array.isArray(userBasedPerms?.groupBased)
      ? (userBasedPerms!.groupBased as GroupRule[])
      : [];

    let disabledByPerms = true;   // default: disabled unless an enabling rule applies
    // let hiddenByPerms = false; // future: default visible

    for (const rule of rules) {
      const src = rule?.groupSource;
      if (!src) continue;

      const inGroup = userGroupTitles.has(norm(src.groupName));
      if (!inGroup) {
        // if user not in this group, this rule does not grant permissions
        continue;
      }

      const editableArr = Array.isArray(src.editableOnly) ? src.editableOnly : [];
      const visibleArr  = Array.isArray(src.visibleOnly)  ? src.visibleOnly  : [];

      const editable = editableArr.some((n) => norm(n) === fieldNameNorm);
      const _visible = visibleArr.some((n) => norm(n) === fieldNameNorm); // future use
      void _visible; // mark as intentionally unused to silence TS warning

      // If any matched group lists this field as editable → enabled
      if (editable) {
        disabledByPerms = false;
      }

      // Future hiding support (leave commented until you enable it):
      // if (_visible) {
      //   hiddenByPerms = false;
      // } else {
      //   hiddenByPerms = true;
      // }
    }

    // // AllHiddenFields future precedence (example):
    // const hiddenByAllHidden =
    //   Array.isArray(AllHiddenFields) &&
    //   AllHiddenFields.some((n) => norm(n) === fieldNameNorm));
    // setIsHidden(hiddenByAllHidden || hiddenByPerms);

    const finalDisabled = !!disabledProp || disabledByPerms;
    setIsDisabled(finalDisabled);
  }, [displayName, disabledProp, AllDisabledFields, AllHiddenFields, userBasedPerms, curUserInfo]);

  // Selection helper (TS-safe)
  const getSelection = (el: HTMLInputElement) => {
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? el.value.length;
    return { start, end };
  };

  // TEXT paste
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

  // NUMBER paste (enforce decimal limit)
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

  // Change
  const handleChange: React.ComponentProps<typeof Input>['onChange'] = (_e, data) => {
    const raw = data.value ?? '';
    const sanitized0 = isNumber ? decimalSanitizer(raw) : raw;
    const { value: next, trimmed } = isNumber
      ? enforceDecimalLimit(sanitized0)
      : { value: sanitized0, trimmed: false };

    setLocalVal(next);

    if (!isNumber && isDefined(maxLength) && next.length >= maxLength) {
      setError(lengthMsg);
      return;
    }

    if (isNumber && trimmed && decimalLimit !== null) {
      setError(decimalLimitMsg(decimalLimit));
      return;
    }

    const nextErr = isNumber ? validateNumber(next) : (touched ? validateText(next) : '');
    setError(nextErr);

    // commitValue(next, nextErr); // opt-in for live commit
  };

  // Blur
  const handleBlur: React.FocusEventHandler<HTMLInputElement> = () => {
    setTouched(true);
    const err =
      (!isNumber && isDefined(maxLength) && localVal.length >= maxLength)
        ? lengthMsg
        : computeError(localVal);
    setError(err);
    commitValue(localVal, err);
  };

  const after = isNumber && contentAfter === 'percentage'
    ? <Text size={400} id={`${inputId}Per`}>%</Text>
    : undefined;

  const hasError = error !== '';
  const stepAttr = isNumber
    ? (decimalLimit === 1 ? '0.1' : decimalLimit === 2 ? '0.01' : 'any')
    : undefined;

  // if (isHidden) return null; // enable when you wire up hiding

  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
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
        // TEXT ONLY
        maxLength={!isNumber && isDefined(maxLength) ? maxLength : undefined}
        // NUMBER ONLY
        type={isNumber ? 'number' : 'text'}
        inputMode={isNumber ? 'decimal' : undefined}
        step={stepAttr}
        min={isNumber && isDefined(min) ? min : undefined}
        max={isNumber && isDefined(max) ? max : undefined}
        contentAfter={after}
      />

      {description !== '' && <div className="descriptionText">{description}</div>}
    </Field>
  );
}
