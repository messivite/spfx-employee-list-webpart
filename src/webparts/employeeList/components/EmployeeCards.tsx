import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClient as SPHttpClientType } from '@microsoft/sp-http';
import { SpErrorBoundary } from '@mustafaaksoy41/sharepoint-kit/components';
import {
  Theme, Text, Spinner, Callout, Flex, Separator,
  DropdownMenu, IconButton, Box, Select, Tooltip
} from '@radix-ui/themes';
import { Avatar } from './ui/Avatar/Avatar';
import {
  MagnifyingGlassIcon, Cross2Icon, DotsHorizontalIcon, PersonIcon
} from '@radix-ui/react-icons';
import { buildGraphUrl } from '../utils/graphClient';
import { EmployeeProfileModal } from './EmployeeProfileModal';
import type { GraphUserProfile } from './EmployeeProfileModal';
import { Employee } from '../../../models/sp-types';

const GUID_REGEX = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

export interface EmployeeCardsProps {
  listId: string;
  recordCount: number;
  isDarkTheme: boolean;
  getAccessToken: () => Promise<string>;
  spHttpClient?: SPHttpClientType;
  webAbsoluteUrl?: string;
}

interface EmployeeItem {
  id: string;
  fields: Employee & {
    id?: string;
    JobTitle?: unknown;
    PersonLookupId?: string | number;
    personDisplayName?: string;
    personEmail?: string;
    personPhotoUrl?: string;
  };
}

interface UserInfo {
  title: string;
  email?: string;
  photoUrl?: string;
}

export function EmployeeCards(props: EmployeeCardsProps): React.ReactElement {
  const { listId, recordCount, getAccessToken, spHttpClient, webAbsoluteUrl } = props;
  const [resolvedListId, setResolvedListId] = React.useState<string | null>(GUID_REGEX.test(listId) ? listId : null);
  const [data, setData] = React.useState<EmployeeItem[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<Error | null>(null);
  const [search, setSearch] = React.useState('');
  const [jobTitleFilter, setJobTitleFilter] = React.useState('all');
  const [profileOpen, setProfileOpen] = React.useState(false);
  const [profile, setProfile] = React.useState<GraphUserProfile | null>(null);
  const [profileLoading, setProfileLoading] = React.useState(false);
  const [profilePhoto, setProfilePhoto] = React.useState<string | undefined>(undefined);
  const userInfoMapRef = React.useRef<Record<string, UserInfo>>({});

  /* ── data fetch ── */
  React.useEffect(() => {
    if (!getAccessToken || !spHttpClient || !webAbsoluteUrl) {
      setLoading(false);
      setError(new Error('SharePoint/Graph bağlantısı gerekli'));
      return;
    }
    const effectiveId = resolvedListId ?? listId;
    if (!GUID_REGEX.test(effectiveId)) {
      const titleUrl = `${webAbsoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(effectiveId)}')?$select=Id`;
      spHttpClient
        .get(titleUrl, SPHttpClient.configurations.v1)
        .then((r) => (r.ok ? r.json() : Promise.reject(new Error('Liste bulunamadı'))))
        .then((list) => { setResolvedListId(list.Id); })
        .catch((err) => { setError(err); setLoading(false); });
      return;
    }
    const rootOrigin = new URL(webAbsoluteUrl).origin;
    const top = Math.min(recordCount, 500);
    setLoading(true);
    setError(null);
    getAccessToken()
      .then((token) => {
        const url = buildGraphUrl(
          `/sites/root/lists/${encodeURIComponent(effectiveId)}/items`,
          `expand=fields&$top=${top}`
        );
        return fetch(url, {
          headers: { Authorization: `Bearer ${token}` }
        });
      })
      .then((res) => {
        if (!res.ok) return res.text().then((t) => Promise.reject(new Error(`Graph ${res.status}: ${t.substring(0, 150)}`)));
        return res.json();
      })
      .then(async (graphBody: { value?: Array<{ id: string; fields?: Record<string, unknown> }> }) => {
        const rawItems = graphBody.value ?? [];
        const uniqueIds = Array.from(new Set(rawItems.map((i) => i.fields?.PersonLookupId).filter(Boolean))) as string[];
        const infoMap: Record<string, UserInfo> = {};

        await Promise.all(
          uniqueIds.map(async (lid) => {
            try {
              const userRes = await spHttpClient!.get(
                `${rootOrigin}/_api/web/siteusers/getbyid(${lid})?$select=Title,Email,LoginName`,
                SPHttpClient.configurations.v1
              );
              if (!userRes.ok) throw new Error('fail');
              const u = await userRes.json();
              const rawEmail = (u.Email as string) || '';
              const loginName = (u.LoginName as string) || '';
              const accountName = (rawEmail || '').trim() || (loginName || '').trim() || undefined;
              const email = rawEmail || undefined;
              const photoUrl = accountName
                ? `${rootOrigin}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(accountName)}`
                : undefined;
              // Debug: hangi kullanıcı için hangi fotoğraf URL'si üretildi?
              // eslint-disable-next-line no-console
              console.log('[EmployeeCards] list user photo', {
                lid,
                title: u.Title,
                email,
                loginName,
                accountName,
                photoUrl
              });
              infoMap[lid] = { title: (u.Title as string) || `ID: ${lid}`, email, photoUrl };
            } catch {
              infoMap[lid] = { title: `ID: ${lid}` };
            }
          })
        );

        userInfoMapRef.current = infoMap;
        const items: EmployeeItem[] = rawItems.map((i) => {
          const fid = String(i.id);
          const f = i.fields ?? {};
          const plid = f.PersonLookupId as string | undefined;
          const info = plid ? infoMap[plid] : undefined;
          return {
            id: fid,
            fields: {
              id: fid,
              JobTitle: f.JobTitle,
              PersonLookupId: plid,
              personDisplayName: info?.title,
              personEmail: info?.email,
              personPhotoUrl: info?.photoUrl
            } as EmployeeItem['fields']
          };
        });
        setData(items);
      })
      .catch((err: Error) => setError(err))
      .then(() => setLoading(false), () => setLoading(false));
  }, [listId, recordCount, spHttpClient, webAbsoluteUrl, resolvedListId, getAccessToken]);

  /* ── helpers ── */
  const getName = (item: EmployeeItem): string => String(item.fields.personDisplayName ?? `ID: ${item.id}`);
  const getEmail = (item: EmployeeItem): string => String(item.fields.personEmail ?? '');
  const getJobTitle = (item: EmployeeItem): string => String(item.fields.JobTitle ?? '');
  const getPhotoUrl = (item: EmployeeItem): string | undefined => item.fields.personPhotoUrl as string | undefined;

  /* ── unique job titles for filter ── */
  const uniqueJobTitles = React.useMemo(() => {
    const titles = data
      .map(item => String(item.fields.JobTitle ?? ''))
      .filter(t => t.length > 0);
    return Array.from(new Set(titles)).sort((a, b) => a.localeCompare(b, 'tr'));
  }, [data]);

  /* ── search + filter ── */
  const matchesSearch = (item: EmployeeItem): boolean => {
    if (!search.trim()) return true;
    const q = search.toLowerCase();
    const name = String(item.fields.personDisplayName ?? '').toLowerCase();
    const job = String(item.fields.JobTitle ?? '').toLowerCase();
    const email = String(item.fields.personEmail ?? '').toLowerCase();
    return name.includes(q) || job.includes(q) || email.includes(q);
  };
  const filteredBySearch = data.filter(matchesSearch);
  const filteredData = jobTitleFilter === 'all'
    ? filteredBySearch
    : filteredBySearch.filter(item => getJobTitle(item) === jobTitleFilter);

  /* ── profile fetch ── */
  const fetchUserProfile = React.useCallback(
    async (personLookupId: number | string): Promise<void> => {
      if (!webAbsoluteUrl || !spHttpClient || !getAccessToken) return;
      setProfileLoading(true);
      setProfile(null);
      setProfilePhoto(undefined);
      try {
        const id = typeof personLookupId === 'string' ? parseInt(personLookupId, 10) : personLookupId;
        const rootOrigin = new URL(webAbsoluteUrl).origin;
        const userRes = await spHttpClient.get(
          `${rootOrigin}/_api/web/siteusers/getbyid(${id})?$select=LoginName,Email`,
          SPHttpClient.configurations.v1
        );
        if (!userRes.ok) throw new Error('Kullanıcı bilgisi alınamadı');
        const userData = await userRes.json();
        let upn = userData.Email as string | undefined;
        if (!upn) {
          const loginName = userData.LoginName as string;
          upn = loginName?.includes('|') ? loginName.split('|').pop()?.trim() : loginName;
        }
        if (!upn) throw new Error('UPN bulunamadı');
        const token = await getAccessToken();
        const userUrl = buildGraphUrl(
          `/users/${encodeURIComponent(upn)}`,
          '$select=displayName,mail,jobTitle,mobilePhone,officeLocation,userPrincipalName,department'
        );
        const graphRes = await fetch(userUrl, {
          headers: { Authorization: `Bearer ${token}` }
        });
        if (!graphRes.ok) throw new Error(`Graph: ${graphRes.status}`);
        const profileData = (await graphRes.json()) as GraphUserProfile;
        setProfile(profileData);
        const profileRootOrigin = new URL(webAbsoluteUrl).origin;
        const profileUrl = `${profileRootOrigin}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(upn)}`;
        // Debug: dialog için profil foto URL'si
        // eslint-disable-next-line no-console
        console.log('[EmployeeCards] profile dialog photo', { upn, profileUrl });
        setProfilePhoto(profileUrl);
        setProfileOpen(true);
      } catch (err) {
        setProfile({ displayName: (err as Error).message } as GraphUserProfile);
        setProfileOpen(true);
      } finally {
        setProfileLoading(false);
      }
    },
    [webAbsoluteUrl, spHttpClient, getAccessToken]
  );

  const handleRowClick = (item: EmployeeItem): void => {
    const lid = item.fields.PersonLookupId;
    if (typeof lid === 'string' || typeof lid === 'number') {
      fetchUserProfile(lid).catch(() => { /* handled */ });
    } else {
      setProfile({ displayName: 'Profil bilgisi yok' } as GraphUserProfile);
      setProfileOpen(true);
    }
  };

  /* ── render ── */
  if (loading) {
    return (
      <Theme appearance="light" accentColor="violet" radius="large" hasBackground={false} style={{ background: 'transparent', minHeight: 'auto' }}>
        <Flex align="center" justify="center" py="8"><Spinner size="3" /></Flex>
      </Theme>
    );
  }
  if (error) {
    return (
      <Theme appearance="light" accentColor="violet" hasBackground={false} style={{ background: 'transparent', minHeight: 'auto' }}>
        <Callout.Root color="red"><Callout.Text>{error.message}</Callout.Text></Callout.Root>
      </Theme>
    );
  }

  return (
    <Theme appearance="light" accentColor="violet" radius="large" hasBackground={false} style={{ background: 'transparent', minHeight: 'auto' }}>
      <SpErrorBoundary onAuthError={() => alert('Graph API Yetki Hatası!')}>
        <Flex direction="column" gap="5">

          {/* ── SEARCH + FILTER ── */}
          <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 16 }}>
            <div style={{ position: 'relative', flex: 1 }}>
              <div style={{
                position: 'absolute',
                left: 16,
                top: '50%',
                transform: 'translateY(-50%)',
                color: 'rgba(148,163,184,1)',
                display: 'flex',
                alignItems: 'center',
                pointerEvents: 'none'
              }}>
                <MagnifyingGlassIcon width="20" height="20" />
              </div>
              <input
                type="text"
                placeholder="Ekip üyesi ara..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                style={{
                  width: '100%',
                  height: 48,
                  padding: '0 16px 0 46px',
                  border: '1px solid rgba(148,163,184,0.35)',
                  borderRadius: 10,
                  background: '#ffffff',
                  color: '#0f172a',
                  fontSize: 15,
                  outline: 'none',
                  boxSizing: 'border-box',
                  transition: 'border-color 0.2s, background 0.2s, box-shadow 0.2s'
                }}
                onFocus={(e) => {
                  e.currentTarget.style.borderColor = 'rgba(88,80,236,0.9)';
                  e.currentTarget.style.boxShadow = '0 0 0 1px rgba(88,80,236,0.4)';
                }}
                onBlur={(e) => {
                  e.currentTarget.style.borderColor = 'rgba(148,163,184,0.35)';
                  e.currentTarget.style.boxShadow = 'none';
                }}
              />
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                <Text size="1" style={{ color: '#6b7280' }}>
                  Pozisyon
                </Text>
                <Select.Root value={jobTitleFilter} onValueChange={setJobTitleFilter} size="2">
                  <Select.Trigger placeholder="Tüm pozisyonlar" radius="large" color="gray" />
                  <Select.Content position="popper" sideOffset={4}>
                    <Select.Item value="all">Tüm pozisyonlar</Select.Item>
                    {uniqueJobTitles.length > 0 && <Select.Separator />}
                    {uniqueJobTitles.map((title) => (
                      <Select.Item key={title} value={title}>
                        {title}
                      </Select.Item>
                    ))}
                  </Select.Content>
                </Select.Root>
              </div>
              {jobTitleFilter !== 'all' && (
                <Tooltip content="Filtreyi temizle">
                  <IconButton
                    variant="ghost"
                    color="gray"
                    size="1"
                    radius="full"
                    onClick={() => setJobTitleFilter('all')}
                  >
                    <Cross2Icon width="14" height="14" />
                  </IconButton>
                </Tooltip>
              )}
            </div>
          </div>

          {/* ── HEADER ── */}
          <div style={{
            display: 'grid',
            gridTemplateColumns: '36px 44px 1fr 1fr 72px',
            gap: 16,
            alignItems: 'center',
            padding: '0 20px 8px',
            borderBottom: '1px solid rgba(255,255,255,0.06)'
          }}>
            <Text size="1" style={{ color: 'rgba(148,163,184,1)', fontWeight: 600 }}>#</Text>
            <span />
            <Text size="1" style={{ color: 'rgba(148,163,184,1)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' }}>İsim</Text>
            <Text size="1" style={{ color: 'rgba(148,163,184,1)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' }}>Pozisyon</Text>
            <span />
          </div>

          {/* ── LIST ── */}
          {filteredData.length === 0 ? (
            <Flex align="center" justify="center" py="8">
              <Text style={{ color: 'rgba(255,255,255,0.4)' }} size="3">Sonuç bulunamadı</Text>
            </Flex>
          ) : (
            <Box
              style={{
                borderRadius: 14,
                overflowY: 'auto',
                maxHeight: 450,
                border: '1px solid rgba(226,232,240,1)',
                background: '#ffffff'
              }}
            >
              {filteredData.map((item, idx) => (
                <React.Fragment key={item.id}>
                  <div
                    role="button"
                    tabIndex={0}
                    onClick={() => handleRowClick(item)}
                    onKeyDown={(e) => { if (e.key === 'Enter') handleRowClick(item); }}
                    style={{
                      display: 'grid',
                      gridTemplateColumns: '36px 44px 1fr 1fr 72px',
                      gap: 16,
                      alignItems: 'center',
                      padding: '14px 20px',
                      cursor: 'pointer',
                      transition: 'background 0.15s',
                      background: 'transparent'
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.background = 'rgba(139,92,246,0.06)'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.background = 'transparent'; }}
                  >
                    <Text size="2" style={{ color: '#94a3b8', fontWeight: 500, fontVariantNumeric: 'tabular-nums' }}>
                      {idx + 1}
                    </Text>

                    <Avatar name={getName(item)} size={40} imageUrl={getPhotoUrl(item)} />

                    <Text size="2" weight="medium" style={{ color: '#0f172a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {getName(item)}
                    </Text>

                    <Text size="2" style={{ color: 'rgba(100,116,139,1)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {getJobTitle(item) || '—'}
                    </Text>

                    <Box style={{ display: 'flex', justifyContent: 'center', gap: 4, alignItems: 'center' }}>
                      <Tooltip content="Profili görüntüle">
                        <IconButton
                          variant="ghost"
                          color="violet"
                          size="1"
                          radius="full"
                          onClick={(e: React.MouseEvent) => { e.stopPropagation(); handleRowClick(item); }}
                        >
                          <PersonIcon width="16" height="16" />
                        </IconButton>
                      </Tooltip>
                      <DropdownMenu.Root>
                        <DropdownMenu.Trigger>
                          <IconButton
                            variant="ghost"
                            color="gray"
                            size="1"
                            radius="full"
                            onClick={(e: React.MouseEvent) => e.stopPropagation()}
                            style={{ color: '#94a3b8' }}
                          >
                            <DotsHorizontalIcon width="16" height="16" />
                          </IconButton>
                        </DropdownMenu.Trigger>
                        <DropdownMenu.Content size="1" variant="soft">
                          <DropdownMenu.Item onClick={() => handleRowClick(item)}>
                            Profili Görüntüle
                          </DropdownMenu.Item>
                          {getEmail(item) && (
                            <DropdownMenu.Item onClick={() => window.open(`mailto:${getEmail(item)}`)}>
                              E-posta Gönder
                            </DropdownMenu.Item>
                          )}
                        </DropdownMenu.Content>
                      </DropdownMenu.Root>
                    </Box>
                  </div>

                  {idx < filteredData.length - 1 && (
                    <Separator size="4" style={{ opacity: 0.05 }} />
                  )}
                </React.Fragment>
              ))}
            </Box>
          )}

          {/* ── FOOTER ── */}
          <Flex justify="end" px="2" style={{ marginTop: 8 }}>
            <Text size="1" style={{ color: '#94a3b8' }}>
              {filteredData.length} / {data.length} kişi gösteriliyor
            </Text>
          </Flex>

          {/* ── PROFILE MODAL ── */}
          <EmployeeProfileModal
            open={profileOpen}
            onOpenChange={(open) => { setProfileOpen(open); if (!open) setProfilePhoto(undefined); }}
            profile={profile}
            profileLoading={profileLoading}
            profilePhoto={profilePhoto}
          />
        </Flex>
      </SpErrorBoundary>
    </Theme>
  );
}

