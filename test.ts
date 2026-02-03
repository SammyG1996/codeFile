const actions = [
  {
    fieldName: "RequestTracker" as InternalFieldNames,
    value: JSON.stringify({
      RqstrDetails,
      Status,
      PrevStatus,
      PrevStepID
    }),
    step: "preSubmit" as ActionStep
  }
];