import * as React from 'react';
import {
  Dialog, Flex, Text, Separator, Spinner, IconButton, Button
} from '@radix-ui/themes';
import { Avatar } from './ui/Avatar/Avatar';
import {
  Cross2Icon, EnvelopeClosedIcon, IdCardIcon, MobileIcon, GlobeIcon
} from '@radix-ui/react-icons';

export interface GraphUserProfile {
  displayName?: string;
  mail?: string;
  jobTitle?: string;
  mobilePhone?: string;
  officeLocation?: string;
  userPrincipalName?: string;
  department?: string;
}

export interface EmployeeProfileModalProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  profile: GraphUserProfile | null;
  profileLoading: boolean;
  profilePhoto?: string;
}

export function EmployeeProfileModal({
  open,
  onOpenChange,
  profile,
  profileLoading,
  profilePhoto
}: EmployeeProfileModalProps): React.ReactElement {
  return (
    <Dialog.Root open={open} onOpenChange={onOpenChange}>
      <Dialog.Content style={{ maxWidth: 440, borderRadius: 16, padding: 0, overflow: 'hidden' }}>
        <div style={{
          height: 100,
          background: 'linear-gradient(135deg, #6d28d9 0%, #7c3aed 50%, #a78bfa 100%)',
          position: 'relative'
        }}>
          <div style={{
            position: 'absolute',
            bottom: -32,
            left: 24,
            borderRadius: '50%',
            border: '4px solid var(--color-background)',
            background: 'var(--color-background)',
            lineHeight: 0
          }}>
            <Avatar name={profile?.displayName ?? '?'} size={64} imageUrl={profilePhoto} />
          </div>
          <Dialog.Close>
            <IconButton variant="ghost" size="1" radius="full" style={{ color: 'rgba(255,255,255,0.8)', position: 'absolute', top: 12, right: 12 }}>
              <Cross2Icon />
            </IconButton>
          </Dialog.Close>
        </div>

        <div style={{ padding: '44px 24px 24px' }}>
          {profileLoading ? (
            <Flex align="center" justify="center" py="6"><Spinner size="3" /></Flex>
          ) : profile ? (
            <Flex direction="column" gap="4">
              <div>
                <Text size="5" weight="bold" style={{ display: 'block' }}>{profile.displayName ?? '-'}</Text>
                {profile.jobTitle && (
                  <Text size="2" color="gray" style={{ display: 'block', marginTop: 2 }}>{profile.jobTitle}</Text>
                )}
              </div>
              <Separator size="4" />
              <Flex direction="column" gap="4">
                <ProfileRow icon={<EnvelopeClosedIcon />} label="E-posta" value={profile.mail ?? profile.userPrincipalName} />
                <ProfileRow icon={<IdCardIcon />} label="Departman" value={profile.department} />
                <ProfileRow icon={<MobileIcon />} label="Telefon" value={profile.mobilePhone} />
                <ProfileRow icon={<GlobeIcon />} label="Konum" value={profile.officeLocation} />
              </Flex>
              <Flex justify="end" mt="5">
                <Dialog.Close>
                  <Button variant="soft" color="gray" style={{ cursor: 'pointer' }}>
                    Kapat
                  </Button>
                </Dialog.Close>
              </Flex>
            </Flex>
          ) : undefined}
        </div>
      </Dialog.Content>
    </Dialog.Root>
  );
}

function ProfileRow({ icon, label, value }: { icon: React.ReactElement; label: string; value?: string }): React.ReactElement {
  if (!value) return <React.Fragment />;
  return (
    <Flex gap="3" align="center">
      <div style={{ color: 'rgba(255,255,255,0.35)', display: 'flex', flexShrink: 0 }}>{icon}</div>
      <Text size="1" color="gray" style={{ minWidth: 72, flexShrink: 0 }}>{label}</Text>
      <Text size="2" style={{ wordBreak: 'break-word', whiteSpace: 'pre-wrap' }}>{value}</Text>
    </Flex>
  );
}
