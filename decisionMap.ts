/** statusText must match exactly to SharePoint Status column choices */
import type { DecisionMap } from './types';

export const decisionMap: DecisionMap = {
  steps: {
    'P200': {
      decisionStepId: 'P200',
      Yes: [
        { id: 'P300', statusText: 'Route for Info' },
        { id: 'P400', statusText: 'Assign to TW' },
        { id: 'P500', statusText: 'Cancelled' }
      ],
      No: 'P600'
    },

    'P800': {
      decisionStepId: 'P800',
      Yes: [
        { id: 'P900', statusText: 'Route for Info' },
        { id: 'P1000', statusText: 'In Progress' },
        { id: 'P1100', statusText: 'Reassign to TW' },
        { id: 'P1200', statusText: 'Completed' },
        { id: 'P1300', statusText: 'Cancelled' }
      ],
      No: 'P1400'
    },

    'P1600': {
      decisionStepId: 'P1600',
      Yes: [
        { id: 'P1700', statusText: 'Route for Info' },
        { id: 'P1800', statusText: 'Reassign to TW' },
        { id: 'P1900', statusText: 'Complete' },
        { id: 'P2000', statusText: 'Cancelled' }
      ],
      No: 'P2100'
    }
  }
};