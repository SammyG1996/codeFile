// RequestTypeSelector.tsx
//
// Renders an instruction line plus a labeled "Request Type" dropdown.
// Title and label text are configurable via props.

import * as React from 'react';
import {
  Select,
  Text,
  Label,
} from '@fluentui/react-components';

export interface RequestTypeSelectorProps {
  id: string;
  requestTypes?: string[];
  value?: string;
  onChange?: (requestType: string) => void;
  title?: string;
  label?: string;
}

const defaultRequestTypes: string[] = [
  'New-Online Help',
  'New-Policy',
  'New-Standard Operating Procedure (SOP)',
  'New-Desktop Procedure (DP)',
  'Change-Online Help',
  'Change-Policy',
  'Change-Standard Operating Procedure (SOP)',
  'Change-Desktop Procedure (DP)',
  'Archive (Online Help Only)',
  'Retire (Procedural Documents Only)',
  'Request for Information-Online Help',
  'Request for Information-Procedural Documents',
];

const RequestTypeSelector = (props: RequestTypeSelectorProps): JSX.Element => {
  const {
    id,
    requestTypes = defaultRequestTypes,
    value,
    onChange,
    title = 'Please select Request Type to begin',
    label = 'Request Type',
  } = props;

  // Local value for the Select
  const [selectedType, setSelectedType] = React.useState<string>(value ?? '');
  // After the first real selection, prevent returning to the blank option
  const [placeholderLocked, setPlaceholderLocked] = React.useState<boolean>(
    Boolean(value)
  );

  // Keep local state in sync if parent changes `value`
  React.useEffect((): void => {
    if (value !== undefined && value !== selectedType) {
      setSelectedType(value);
      setPlaceholderLocked(Boolean(value));
    }
  }, [value, selectedType]);

  /**
   * Native select-style change handler.
   * This is just a regular ChangeEvent on an HTMLSelectElement.
   */
  const handleChange: React.ChangeEventHandler<HTMLSelectElement> = (
    event
  ): void => {
    const next = event.target.value ?? '';
    setSelectedType(next);
    if (next) setPlaceholderLocked(true);
    if (onChange) onChange(next);
  };

  return (
    <div
      className="ks-requestTypeWrapper"
    >
      {/* Instruction above the field */}
      <div style={{ marginBottom: 12 }}>
        <Text
          as="p"
          size={400}
          weight="semibold"
          style={{
            color: '#c00000',
            margin: 0,
            letterSpacing: '-0.2px',
          }}
        >
          {title}
        </Text>
      </div>

      {/* Request Type row */}
      <div className="ks-requestTypeRow">
        <Label
          htmlFor={id}
          style={{
            display: 'flex',
            alignItems: 'baseline',
            gap: 4,
            fontWeight: 600,
            fontSize: 18,
            color: '#444',
            marginBottom: 6,
          }}
        >
          <span>{label}</span>
          <span style={{ color: '#c00000' }}>*</span>
        </Label>

        <Select
          id={id}
          name={id}
          value={selectedType}
          onChange={handleChange}
          style={{
            width: '100%',
            border: '1px solid #9a9a9a',
            borderRadius: 0,
            backgroundColor: '#fff',
            padding: '10px 10px',
            fontSize: 16,
            color: '#333',
            boxShadow: 'none',
            outline: 'none',
            minHeight: 42,
          }}
          title={selectedType || 'Select a request type'}
        >
          {/* Placeholder keeps the field blank initially */}
          <option
            value=""
            disabled={placeholderLocked}
            aria-label="Select a request type"
          />

          {requestTypes.map((rt) => (
            <option key={rt} value={rt}>
              {rt}
            </option>
          ))}
        </Select>
      </div>
    </div>
  );
};

export default RequestTypeSelector;
