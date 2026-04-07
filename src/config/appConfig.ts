// Configuration constants for the QQIA Agent
export const config = {
  bot: {
    port: parseInt(process.env.PORT || '3978'),
    appId: process.env.MICROSOFT_APP_ID || '',
    appPassword: process.env.MICROSOFT_APP_PASSWORD || '',
    tenantId: process.env.MICROSOFT_APP_TENANT_ID || '',
  },
  cosmos: {
    endpoint: process.env.COSMOS_ENDPOINT || '',
    key: process.env.COSMOS_KEY || '',
    database: process.env.COSMOS_DATABASE || 'qqia-agent',
    containers: {
      steps: 'steps',
      milestones: 'milestones',
      audit: 'audit',
      users: 'users',
    },
  },
  graph: {
    clientId: process.env.GRAPH_CLIENT_ID || '',
    clientSecret: process.env.GRAPH_CLIENT_SECRET || '',
    tenantId: process.env.GRAPH_TENANT_ID || '',
    excelDriveId: process.env.EXCEL_DRIVE_ID || '',
    excelItemId: process.env.EXCEL_ITEM_ID || '',
    excelFilePath: process.env.EXCEL_FILE_PATH || '',
  },
  notifications: {
    channelId: process.env.NOTIFICATION_CHANNEL_ID || '',
    teamId: process.env.NOTIFICATION_TEAM_ID || '',
    deadlineWarningDays: [3, 1],
    overdueEscalationDays: 3,
    syncIntervalMinutes: 15,
    weeklyDigestCron: '0 8 * * 1', // Monday 8 AM
  },
  statuses: ['Not Started', 'In Progress', 'Completed', 'Blocked', 'N/A'] as const,
  tracks: ['Corp', 'Fed'] as const,
};

export type StepStatus = typeof config.statuses[number];
export type Track = typeof config.tracks[number];
