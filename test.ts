/**these are the cases */

case "Assigned to TW": {
  if (!ctx.twEmail || !ctx.twName) return null;

  const editUrl = `${Config.editPageBaseUrl}${ctx.itemID}`;

  return {
    to: [ctx.twEmail],
    subject: `Let's Fix It: ${ctx.requestTypeText ?? "Request"} (${ctx.itemID}) assigned to you`,
    html: renderAssignedToTW({
      twName: ctx.twName,
      requesterName: ctx.requesterName ?? "",
      requestTypeText: ctx.requestTypeText ?? "",
      itemID: String(ctx.itemID),
      editUrl,
    }),
  };
}

case "Route to Router":
case "Routed": {
  if (!ctx.routerEmail || !ctx.routerName) return null;

  const editUrl = `${Config.editPageBaseUrl}${ctx.itemID}`;

  return {
    to: [ctx.routerEmail],
    subject: `Knowledge Services Request for: Request ${ctx.itemID} ${ctx.status} - no reply`,
    html: renderRoutedToRouter({
      routerName: ctx.routerName,
      requestTypeText: ctx.requestTypeText ?? "",
      requesterName: ctx.requesterName ?? "",
      itemID: String(ctx.itemID),
      editUrl,
      businessStatus: ctx.businessStatus,
      additionalInfo: ctx.additionalInfo,
    }),
  };
}


/**this goes after the cases  */

/**
 * Builds additional emails for statuses that require MORE than one recipient.
 * This mirrors the legacy JS behavior (requester + internal mailbox, etc.).
 *
 * Call this AFTER buildEmail(ctx) if needed.
 */
export function buildAdditionalEmails(ctx: RouterContext): EmailPayload[] {
  const editUrl = `${Config.editPageBaseUrl}${ctx.itemID}`;
  const viewUrl = `${Config.viewPageBaseUrl}${ctx.itemID}`;

  switch (ctx.status) {
    case "Received":
    case "Submitted": {
      if (!ctx.requesterEmail || !ctx.requesterName) return [];

      return [
        // requester confirmation
        {
          to: [ctx.requesterEmail],
          subject: `Your Knowledge Services Let's Fix It request regarding ${
            ctx.requestTypeText ?? "your request"
          } has been received`,
          html: renderRequesterReceived({
            requesterName: ctx.requesterName,
            requestTypeText: ctx.requestTypeText ?? "",
            itemID: String(ctx.itemID),
          }),
        },

        // internal mailbox notification
        {
          to: [Config.internalMailbox],
          subject: `New Let's Fix It request for ${ctx.requestTypeText ?? "Request"} has been submitted by ${
            ctx.requesterName ?? "Requester"
          } (${ctx.itemID})`,
          html: renderInternalNewRequest({
            requestTypeText: ctx.requestTypeText ?? "",
            requesterName: ctx.requesterName,
            itemID: String(ctx.itemID),
            editUrl,
          }),
        },
      ];
    }

    case "Returned from Route":
    case "Returned": {
      if (!ctx.requesterEmail || !ctx.requesterName) return [];

      const statusText = ctx.businessStatus ?? ctx.status;

      return [
        // requester notification
        {
          to: [ctx.requesterEmail],
          subject: `Knowledge Services Request ${ctx.itemID} has been returned from business`,
          html: renderReturnedToRequester({
            requesterName: ctx.requesterName,
            requestTypeText: ctx.requestTypeText ?? "",
            itemID: String(ctx.itemID),
            statusText,
            viewUrl,
          }),
        },

        // internal mailbox copy
        {
          to: [Config.internalMailbox],
          subject: `Knowledge Services Let's Fix It Form: Request ${ctx.itemID} Returned from Route - no reply`,
          html: renderReturnedInternal({
            requestTypeText: ctx.requestTypeText ?? "",
            itemID: String(ctx.itemID),
            statusText,
            viewUrl,
          }),
        },
      ];
    }

    default:
      return [];
  }
}


/**How these two pieces work together (no code changes required) */

const primary = buildEmail(ctx);
if (primary) await sendEmailViaSharePoint(sp, primary);

for (const extra of buildAdditionalEmails(ctx)) {
  await sendEmailViaSharePoint(sp, extra);
}
