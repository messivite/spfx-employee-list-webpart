import * as React from 'react';
import * as AvatarPrimitive from '@radix-ui/react-avatar';

interface AvatarProps {
    name: string;
    imageUrl?: string;
    size?: number;
}

function getInitials(str: string): string {
    const parts = (str || '').split(' ').filter(Boolean);
    if (parts.length === 0) return '?';
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function getAvatarColor(str: string): string {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    const hue = Math.abs(hash) % 360;
    return `hsl(${hue}, 55%, 50%)`;
}

export function Avatar({ name, imageUrl, size = 40 }: AvatarProps): React.ReactElement {
    const bgColor = getAvatarColor(name);
    const fontSize = Math.max(11, size * 0.38);

    return (
        <AvatarPrimitive.Root
            style={{
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                verticalAlign: 'middle',
                overflow: 'hidden',
                userSelect: 'none',
                width: size,
                height: size,
                minWidth: size,
                borderRadius: '50%',
                backgroundColor: bgColor,
                flexShrink: 0
            }}
        >
            {imageUrl && (
                <AvatarPrimitive.Image
                    style={{ width: '100%', height: '100%', objectFit: 'cover', borderRadius: '50%' }}
                    src={imageUrl}
                    alt={name}
                />
            )}
            <AvatarPrimitive.Fallback
                style={{
                    width: '100%',
                    height: '100%',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    color: '#fff',
                    fontSize,
                    fontWeight: 600,
                    letterSpacing: '0.5px',
                    lineHeight: 1
                }}
                delayMs={imageUrl ? 400 : 0}
            >
                {getInitials(name)}
            </AvatarPrimitive.Fallback>
        </AvatarPrimitive.Root>
    );
}
