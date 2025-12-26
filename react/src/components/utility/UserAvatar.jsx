import React, { useState, useEffect } from 'react';

/**
 * Generate initials from a name
 */
function getInitials(name) {
  if (!name) return '?';

  const parts = name.trim().split(' ');
  if (parts.length >= 2) {
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
  return name.substring(0, 2).toUpperCase();
}

/**
 * Generate a color based on the name (consistent hash)
 */
function getColorFromName(name) {
  if (!name) return '#9ca3af';

  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }

  const colors = [
    '#3b82f6', // blue
    '#8b5cf6', // purple
    '#ec4899', // pink
    '#f59e0b', // amber
    '#10b981', // green
    '#06b6d4', // cyan
    '#6366f1', // indigo
    '#f97316', // orange
  ];

  return colors[Math.abs(hash) % colors.length];
}

/**
 * UserAvatar component
 * Shows user initials with color-coded background
 *
 * Note: Fetching actual profile photos from Microsoft Graph requires
 * On-Behalf-Of flow with backend service. See GRAPH_API_SETUP.md
 */
export default function UserAvatar({ userName, size = 64 }) {
  const initials = getInitials(userName);
  const backgroundColor = getColorFromName(userName);

  return (
    <div
      style={{
        width: size,
        height: size,
        borderRadius: '50%',
        backgroundColor,
        color: 'white',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontSize: size * 0.4,
        fontWeight: 600,
        userSelect: 'none',
        boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
      }}
      title={userName || 'User'}
      aria-label={`Avatar for ${userName || 'user'}`}
    >
      {initials}
    </div>
  );
}

/**
 * Advanced version that attempts to fetch user photo from Graph API
 * (Currently disabled - requires backend OBO flow)
 */
export function UserAvatarWithPhoto({ userName, accessToken, size = 64 }) {
  const [photoUrl, setPhotoUrl] = useState(null);
  const [loading, setLoading] = useState(false);

  // Note: This is commented out because it requires Graph API access
  // which needs On-Behalf-Of flow with backend service
  /*
  useEffect(() => {
    if (!accessToken) return;

    const fetchPhoto = async () => {
      setLoading(true);
      try {
        // This would need to go through your backend with OBO flow
        const response = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
          headers: {
            'Authorization': `Bearer ${accessToken}`
          }
        });

        if (response.ok) {
          const blob = await response.blob();
          setPhotoUrl(URL.createObjectURL(blob));
        }
      } catch (error) {
        console.log('Could not fetch user photo:', error);
      } finally {
        setLoading(false);
      }
    };

    fetchPhoto();
  }, [accessToken]);
  */

  // For now, always show initials avatar
  return <UserAvatar userName={userName} size={size} />;
}
