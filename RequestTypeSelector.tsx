// RequestTypeSelector.tsx
//
// Renders the "Knowledge Services Request Form" header block plus the
// "Request Type:*" dropdown as a single component.

import * as React from 'react';
import {
  Combobox,
  Option,
  Text,
  Label,
} from '@fluentui/react-components';

export interface RequestTypeSelectorProps {
  id: string;
  requestTypes?: string[];
  value?: string;
  onChange?: (requestType: string) => void;
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
  } = props;

  // Local value for the Combobox
  const [selectedType, setSelectedType] = React.useState<string>(value ?? '');

  // Keep local state in sync if parent changes value
  React.useEffect((): void => {
    if (value !== undefined && value !== selectedType) {
      setSelectedType(value);
    }
  }, [value, selectedType]);

  // Correctly typed handler for option selection
  const handleOptionSelect: React.ComponentProps<typeof Combobox>['onOptionSelect'] =
    (_event, data): void => {
      const next =
        (data.optionValue as string | undefined) ??
        (data.optionText as string | undefined) ??
        '';

      setSelectedType(next);
      if (onChange) onChange(next);
    };

  // Correctly typed handler for text change in the Combobox input
  const handleChange: React.ComponentProps<typeof Combobox>['onChange'] =
    (_event, data): void => {
      const next = data.value ?? '';
      setSelectedType(next);
      if (onChange) onChange(next);
    };

  return (
    <div
      className="ks-requestTypeWrapper"
      style={{
        margin: '24px auto',
        maxWidth: 900,
        border: '1px solid #ddd',
        backgroundColor: '#fafafa',
        padding: '24px 32px 32px',
      }}
    >
      {/* Header block */}
      <div style={{ textAlign: 'center', marginBottom: 24 }}>
        <Text
          size={500}
          weight="semibold"
          style={{ display: 'block', marginBottom: 8 }}
        >
          Knowledge Services Request Form
        </Text>

        <Text
          size={300}
          style={{ display: 'block', marginBottom: 8 }}
        >
          Please use this form to submit Knowledge Services Requests
        </Text>

        <Text
          size={300}
          style={{ display: 'block', color: 'red', fontWeight: 500 }}
        >
          Note: * Red Asterisk indicates a required field
        </Text>
      </div>

      {/* Instruction above the field */}
      <div style={{ marginBottom: 8 }}>
        <Text
          size={300}
          style={{ color: 'red', fontWeight: 500 }}
        >
          Please select Request Type to begin
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
        <div style={{ minWidth: 120 }}>
          <Label
            htmlFor={id}
            required
            style={{ fontWeight: 500 }}
          >
            Request Type:
          </Label>
        </div>

        {/* Dropdown column */}
        <div style={{ flex: 1 }}>
          <Combobox
            id={id}
            appearance="outline"
            placeholder="Select a request type"
            value={selectedType}
            onChange={handleChange}
            onOptionSelect={handleOptionSelect}
            title={selectedType || 'Select a request type'}
            aria-label="Request Type"
          >
            {requestTypes.map((rt) => (
              <Option key={rt} value={rt}>
                {rt}
              </Option>
            ))}
          </Combobox>
        </div>
      </div>
    </div>
  );
};

export default RequestTypeSelector;
