export interface IEmployeeListProps {
  description: string;
  listNameOrId: string;
  recordCount: number;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  getAccessToken: () => Promise<string>;
  /** SharePoint REST için: Person kolonunu expand edip isim göstermek için kullanılır */
  spHttpClient?: import('@microsoft/sp-http').SPHttpClient;
  webAbsoluteUrl?: string;
}
