import * as React from 'react';
import '@radix-ui/themes/styles.css';
import styles from './EmployeeList.module.scss';
import type { IEmployeeListProps } from './IEmployeeListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SpProvider, SpErrorBoundary } from '@mustafaaksoy41/sharepoint-kit/components';
import { motion } from 'framer-motion';
import { EmployeeCards } from './EmployeeCards';
import { PersonIcon } from '@radix-ui/react-icons';
import * as strings from 'EmployeeListWebPartStrings';

const DEFAULT_LIST_ID = '7b906718-e108-42ac-92df-999e1c6e6e65';

export default class EmployeeList extends React.Component<IEmployeeListProps> {
  public render(): React.ReactElement<IEmployeeListProps> {
    const {
      isDarkTheme,
      userDisplayName,
      getAccessToken,
      spHttpClient,
      webAbsoluteUrl,
      listNameOrId,
      recordCount
    } = this.props;

    return (
      <section className={`${styles.employeeList}`}>
        <SpProvider siteId="root" getAccessToken={getAccessToken}>
          <SpErrorBoundary onAuthError={() => alert("Graph API Yetki Hatası!")}>
            <motion.div
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.4, ease: 'easeOut' }}
              style={{
                fontFamily: "'Segoe UI', Inter, system-ui, -apple-system, sans-serif",
                padding: '32px',
                background: 'linear-gradient(155deg, #fdfdff 0%, #f4f6fb 45%, #ffffff 100%)',
                borderRadius: 20,
                boxShadow: '0 10px 32px rgba(15,23,42,0.12), 0 0 0 1px rgba(226,232,240,0.9)',
                minWidth: 0,
                height: 'max-content'
              }}
            >
              {/* Header */}
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 8 }}>
                <div style={{
                  width: 40, height: 40, borderRadius: 12,
                  background: 'linear-gradient(135deg, #7c3aed, #a78bfa)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  boxShadow: '0 4px 12px rgba(124,58,237,0.3)'
                }}>
                  <PersonIcon width="20" height="20" color="white" />
                </div>
                <div>
                  <h2 style={{
                    margin: 0, fontSize: 20, fontWeight: 700,
                    color: '#0f172a', letterSpacing: '-0.3px', lineHeight: 1.3
                  }}>
                    {strings.TeamMembersTitle}
                  </h2>
                  <p style={{
                    margin: 0, color: '#6b7280', fontSize: 13
                  }}>
                    {strings.WelcomeText}<strong style={{ color: '#111827' }}>{escape(userDisplayName)}</strong>
                  </p>
                </div>
              </div>

              <div style={{ marginTop: 24 }}>
                <EmployeeCards
                  listId={(listNameOrId || DEFAULT_LIST_ID).trim()}
                  recordCount={recordCount ?? 10}
                  isDarkTheme={isDarkTheme}
                  getAccessToken={getAccessToken}
                  spHttpClient={spHttpClient}
                  webAbsoluteUrl={webAbsoluteUrl}
                />
              </div>
            </motion.div>
          </SpErrorBoundary>
        </SpProvider>
      </section>
    );
  }
}
