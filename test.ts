switch (ctx.status) {
  case "Assigned to TW": {
    // existing logic stays exactly as-is
    if (!ctx.twName || !ctx.twEmail) return null;

    const v: AssignedToTWVars = {
      twFullName: ctx.twName,
      requesterFullName: ctx.requesterName,
      requestTypeText: ctx.requestTypeText,
      requestId: String(ctx.itemID),
      editPageUrl: editUrl,
    };

    const subject = subjectAssignedToTw(v);
    const html = renderAssignedToTw(v);

    return { to: [ctx.twEmail], subject, html };
  }

  case "Route to Submitter": {
    break;
  }

  case "Route to Team Lead": {
    break;
  }

  case "Returned from Route": {
    break;
  }

  case "Completed": {
    break;
  }

  case "Cancelled": {
    break;
  }

  default:
    return null;
}