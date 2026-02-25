export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Basit Graph URL helper'ı.
 * Örnek:
 *   buildGraphUrl('/users/UPN', '$select=displayName,mail')
 */
export const buildGraphUrl = (path: string, query?: string): string =>
  `${GRAPH_BASE}${path}${query ? (path.includes('?') ? `&${query}` : `?${query}`) : ''}`;

