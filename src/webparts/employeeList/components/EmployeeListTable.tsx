import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClient as SPHttpClientType } from '@microsoft/sp-http';
import { SpProvider, SpListTable, SpErrorBoundary } from '@mustafaaksoy41/sharepoint-kit/components';
import { Table, Text, Spinner, Callout } from '@radix-ui/themes';
import { Employee } from '../../../models/sp-types';
import { Avatar } from './ui/Avatar/Avatar';

const EMPLOYEE_LIST_ID = '7b906718-e108-42ac-92df-999e1c6e6e65';
/** SharePoint'te Person kolonunun internal adı. Listede farklıysa (örn. Personel) buradan değiştirin. */
const PERSON_FIELD_NAME = 'Person';

export interface EmployeeListTableProps {
  isDarkTheme: boolean;
  userDisplayName: string;
  getAccessToken: () => Promise<string>;
  spHttpClient?: SPHttpClientType;
  webAbsoluteUrl?: string;
}

interface RestListItem {
  ID: number;
  Title?: string;
  JobTitle?: string;
  PersonLookupId?: number;
  [key: string]: unknown;
}

function formatJobTitle(_val: unknown, item: Record<string, unknown>): string {
  return String(item.JobTitle ?? item.jobTitle ?? '-');
}

function renderPersonCell(_value: unknown, item: Record<string, unknown>): React.ReactNode {
  const person = item[PERSON_FIELD_NAME] as { Title?: string } | undefined;
  const title = item.Title ?? item.title;
  const lookupId = item.PersonLookupId ?? item.userId;
  const name = person?.Title ?? title ?? (lookupId !== undefined && lookupId !== null ? `Personel ID: ${lookupId}` : 'Bilinmiyor');
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
      <Avatar name={String(name)} size={36} />
      <span style={{ fontWeight: 600 }}>{String(name)}</span>
    </div>
  );
}

export function EmployeeListTable(props: EmployeeListTableProps): React.ReactElement {
  const { getAccessToken, spHttpClient, webAbsoluteUrl } = props;
  const [restData, setRestData] = React.useState<Array<{ id: string; fields: Record<string, unknown> }> | undefined>(undefined);
  const [loading, setLoading] = React.useState(!!(spHttpClient && webAbsoluteUrl));
  const [error, setError] = React.useState<Error | null>(null);

  React.useEffect(() => {
    if (!spHttpClient || !webAbsoluteUrl) {
      setRestData(undefined);
      setLoading(false);
      return;
    }
    const url = `${webAbsoluteUrl}/_api/web/lists(guid'${EMPLOYEE_LIST_ID}')/items?$select=Id,Title,JobTitle,${PERSON_FIELD_NAME}/Id,${PERSON_FIELD_NAME}/Title&$expand=${PERSON_FIELD_NAME}&$top=5000`;
    setLoading(true);
    setError(null);
    spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((res) => {
        if (!res.ok) {
          throw new Error(`REST ${res.status}: ${res.statusText}`);
        }
        return res.json();
      })
      .then((body: { value?: RestListItem[] }) => {
        const items = body.value ?? [];
        setRestData(
          items.map((item) => ({
            id: String(item.ID),
            fields: {
              id: String(item.ID),
              ID: item.ID,
              Title: item.Title,
              title: item.Title,
              JobTitle: item.JobTitle,
              jobTitle: item.JobTitle,
              [PERSON_FIELD_NAME]: item[PERSON_FIELD_NAME],
              PersonLookupId: (item[PERSON_FIELD_NAME] as { Id?: number })?.Id ?? item.PersonLookupId
            } as Record<string, unknown>
          }))
        );
      })
      .catch((err: Error) => setError(err))
      .then(() => setLoading(false), () => setLoading(false));
  }, [spHttpClient, webAbsoluteUrl]);

  const columns = [
    { key: 'id', label: 'ID' },
    { key: 'title', label: 'Personel', render: renderPersonCell },
    { key: 'jobTitle', label: 'Görev / Unvan', render: (val: unknown, item: Record<string, unknown>) => formatJobTitle(val, item) }
  ];

  if (loading) {
    return <Spinner size="3" />;
  }
  if (error) {
    return (
      <Callout.Root color="red">
        <Callout.Text>
          Liste yüklenemedi: {error.message}. Person alanı için SharePoint REST kullanıldı; yetki veya kolon adı kontrol edin.
        </Callout.Text>
      </Callout.Root>
    );
  }
  if (restData !== undefined) {
    if (restData.length === 0) {
      return <Text color="gray">Kayıt bulunamadı</Text>;
    }
    return (
      <Table.Root>
        <Table.Header>
          <Table.Row>
            {columns.map((col) => (
              <Table.ColumnHeaderCell key={col.key}>{col.label}</Table.ColumnHeaderCell>
            ))}
          </Table.Row>
        </Table.Header>
        <Table.Body>
          {restData.map((item) => (
            <Table.Row key={item.id}>
              {columns.map((col) => (
                <Table.Cell key={col.key}>
                  {col.render ? col.render(item.fields[col.key], item.fields) : String(item.fields[col.key] ?? '')}
                </Table.Cell>
              ))}
            </Table.Row>
          ))}
        </Table.Body>
      </Table.Root>
    );
  }

  return (
    <SpProvider siteId="root" getAccessToken={getAccessToken}>
      <SpErrorBoundary onAuthError={() => alert('Graph API Yetki Hatası!')}>
        <SpListTable<Employee & Record<string, unknown>>
          listId={EMPLOYEE_LIST_ID}
          contentTypeName="Öğe"
          columns={[
            { key: 'id', label: 'ID' },
            { key: 'title', label: 'Personel', render: renderPersonCell },
            { key: 'jobTitle', label: 'Görev / Unvan', render: (val, item) => formatJobTitle(val, item) }
          ]}
        />
      </SpErrorBoundary>
    </SpProvider>
  );
}
