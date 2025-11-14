// RequestTypeSelector.tsx
//
// Renders the top "Knowledge Services Request Form" block plus the
// "Request Type:*" dropdown as a single reusable component.
//
// Parent components can pass in a list of request types and get
// notified when the user picks one. The visual layout is intended
// to closely match the reference screenshot.

import * as React from 'react';
import {
  Combobox,
  Option,
  Text,
  Label,
} from '@fluentui/react-components';

export interface RequestTypeSelectorProps {
  /**
   * Unique id for the dropdown input (used for accessibility and testing).
   */
  id: string;

  /**
   * Available request types to show in the dropdown.
   * If not provided, a default list is used.
   */
  requestTypes?: string[];

  /**
   * Optional preselected value, e.g. when editing.
   */
  value?: string;

  /**
   * Called whenever the user chooses a request type.
   * The callback gets the selected string (or empty string if cleared).
   */
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

/**
 * Top-of-form request type selector.
 * This matches the “Knowledge Services Request Form” header and
 * the “Request Type:*” dropdown area as a single component.
 */
const RequestTypeSelector: React.FC<RequestTypeSelectorProps> = (props) => {
  const {
    id,
    requestTypes = defaultRequestTypes,
    value,
    onChange,
  } = props;

  // Track the current dropdown value locally so the Combobox stays controlled.
  const [selectedType, setSelectedType] = React.useState<string>(value ?? '');

  // Keep local state in sync if the parent changes `value` prop.
  React.useEffect(() => {
    if (value !== undefined && value !== selectedType) {
      setSelectedType(value);
    }
  }, [value, selectedType]);

  /**
   * Fired when the user selects an option from the Combobox list.
   */
  const handleOptionSelect: React.ComponentProps<typeof Combobox>['onOptionSelect'] =
    (_event, data): void => {
      // Prefer optionValue, fall back to optionText, then empty string.
      const next =
        (data.optionValue as string | undefined) ??
        (data.optionText as string | undefined) ??
        '';

      setSelectedType(next);
      onChange?.(next);
    };

  /**
   * Fired when the user types directly into the input.
   * We allow typing, but still treat it as the current value.
   */
  const handleChange: React.ComponentProps<typeof Combobox>['onChange'] =
    (_event, data): void => {
      const next = data.value ?? '';
      setSelectedType(next);
      onChange?.(next);
    };

  return (
    <div
      className="ks-requestTypeWrapper"
      style={{
        // Outer wrapper is a light “card” similar to the screenshot.
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

      {/* Request Type field row */}
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
            // Show a helpful tooltip on hover
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
