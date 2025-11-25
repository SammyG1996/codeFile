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
  // Once a user picks a value, lock the control to prevent re-selection
  const [isLocked, setIsLocked] = React.useState<boolean>(Boolean(value));

  // Keep local state in sync if parent changes `value`
  React.useEffect((): void => {
    if (value !== undefined && value !== selectedType) {
      setSelectedType(value);
      setIsLocked(Boolean(value));
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
    if (next) setIsLocked(true);
    if (onChange) onChange(next);
  };

  return (
    <div
      className="ks-requestTypeWrapper"
      style={{
        margin: '16px 0',
        maxWidth: 1100,
        padding: '8px 4px',
        color: '#4a4a4a',
        fontFamily: '"Times New Roman", Georgia, serif',
      }}
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
      <div
        className="ks-requestTypeRow"
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 8,
        }}
      >
        {/* Label column */}
        <div style={{ minWidth: 150, paddingTop: 2 }}>
          <Label
            htmlFor={id}
            style={{
              fontWeight: 600,
              fontSize: 18,
              color: '#444',
              display: 'flex',
              alignItems: 'baseline',
              gap: 4,
            }}
          >
            <span>{label}</span>
            <span style={{ color: '#c00000' }}>*</span>
          </Label>
        </div>

        {/* Dropdown column */}
        <div style={{ flex: 1 }}>
          <Select
            id={id}
            name={id}
            value={selectedType}
            onChange={handleChange}
            style={{
              width: '100%',
              border: '1px solid #999',
              borderRadius: 2,
              backgroundColor: '#fff',
              padding: '10px 12px',
              fontSize: 15,
              color: '#333',
              boxShadow: 'none',
              outline: 'none',
            }}
            disabled={isLocked}
            title={selectedType || 'Select a request type'}
          >
            {/* Placeholder keeps the field blank initially */}
            <option value="" disabled aria-label="Select a request type" />

            {requestTypes.map((rt) => (
              <option key={rt} value={rt}>
                {rt}
              </option>
            ))}
          </Select>
        </div>
      </div>
    </div>
  );
};

export default RequestTypeSelector;
