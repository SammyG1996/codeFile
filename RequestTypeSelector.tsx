// RequestTypeSelector.tsx
//
// Renders an instruction line plus a labeled dropdown for picking a request type.
// Every behavior is explained inline so a newcomer can follow the flow.

import * as React from 'react';
import {
  Select,
  Text,
  Label,
} from '@fluentui/react-components';

export interface RequestTypeSelectorProps {
  // The control id used for the Select and its associated label
  id: string;
  // Optional list of request type choices; defaults are used if nothing is provided
  requestTypes?: string[];
  // Optional current value controlled by a parent
  value?: string;
  // Optional callback to inform a parent when the selection changes
  onChange?: (requestType: string) => void;
  // Optional heading text shown above the field
  title?: string;
  // Optional label text shown next to the red required asterisk
  label?: string;
}

// Default choices to fall back to when no list is passed in
const defaultRequestTypes: string[] = [
  'New-Online Help',
  'New-Policy',
  'New-Standard Operating Procedure (SOP)',
  'New-Desktop Procedure (DP)'
  // 'Change-Online Help',
  // 'Change-Policy',
  // 'Change-Standard Operating Procedure (SOP)',
  // 'Change-Desktop Procedure (DP)',
  // 'Archive (Online Help Only)',
  // 'Retire (Procedural Documents Only)',
  // 'Request for Information-Online Help',
  // 'Request for Information-Procedural Documents',
];

const RequestTypeSelector = (props: RequestTypeSelectorProps): JSX.Element => {
  // Pull out props and apply defaults so the component works even when props are omitted
  const {
    id,
    requestTypes = defaultRequestTypes,
    value,
    onChange,
    title = 'Please select Request Type to begin',
    label = 'Request Type',
  } = props;

  // Track the current selection inside this component; seed from parent value if provided
  const [selectedType, setSelectedType] = React.useState<string>(value ?? '');
  // Remember if the user has picked any non-empty option; once true, the blank placeholder gets disabled
  const [placeholderLocked, setPlaceholderLocked] = React.useState<boolean>(
    Boolean(value)
  );

  // Keep our local selection in sync if the parent changes the value prop later
  React.useEffect((): void => {
    if (value !== undefined && value !== selectedType) {
      setSelectedType(value);
      setPlaceholderLocked(Boolean(value));
    }
  }, [value, selectedType]);

  // Handle the change event from the native select element
  const handleChange: React.ChangeEventHandler<HTMLSelectElement> = (
    event
  ): void => {
    const next = event.target.value ?? ''; // Read the new selection safely
    setSelectedType(next); // Update local state so the UI reflects the new choice

    // Once a real value is chosen, block the blank placeholder from being reselected
    if (next) setPlaceholderLocked(true);

    // If a parent wants to know about the change, tell it
    if (onChange) onChange(next);
  };

  return (
    <div className="ks-requestTypeWrapper">
      {/* Instruction above the field, shown in red to match the provided reference */}
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

      {/* Request type field: stacked label and select so it aligns with other inputs on the page */}
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
            border: 'none',
            borderRadius: 0,
            backgroundColor: 'transparent',
            padding: 0,
            fontSize: 16,
            color: '#333',
            boxShadow: 'none',
            outline: 'none',
            minHeight: 36,
          }}
          title={selectedType || 'Select a request type'}
        >
          {/* Placeholder keeps the field blank initially; it disables itself after first real choice */}
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
