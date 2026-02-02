/**
 * Single decide() router that delegates to typed decision functions.
 * Use this as the `decide` option in useFlowEngine().
 */
export const decideLetsFixIt: DecisionResolver<LetsFixItFormData> = (step, ctx) => {
  switch (step.id) {
    case 'P200':
      return decideP200(step, ctx);

    case 'P900':
      return decideP900(step, ctx);

    case 'P1800':
      return decideP1800(step, ctx);

    default:
      // If a non-decision step accidentally calls decide, fall back to first outgoing edge.
      // If there is no outgoing edge, end at P2400 (matches the decisionMap placeholder behavior).
      return (step.edges[0]?.to ?? 'P2400') as StepId;
  }
};
