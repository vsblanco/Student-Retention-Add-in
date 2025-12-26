import React, { useState, useEffect } from 'react';

/**
 * Decode JWT token to extract user claims
 */
function decodeAccessToken(token) {
  try {
    const base64Url = token.split('.')[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(
      atob(base64)
        .split('')
        .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
        .join('')
    );
    return JSON.parse(jsonPayload);
  } catch (e) {
    console.error('Failed to decode token:', e);
    return null;
  }
}

export default function UserInfoDisplay({ accessToken }) {
  const [userInfo, setUserInfo] = useState(null);

  useEffect(() => {
    // Try to load cached user info first
    const cachedInfo = localStorage.getItem('SSO_USER_INFO');
    if (cachedInfo) {
      try {
        setUserInfo(JSON.parse(cachedInfo));
      } catch (e) {
        console.error('Failed to parse cached user info:', e);
      }
    }

    // If we have a fresh token, decode and update cache
    if (accessToken) {
      const claims = decodeAccessToken(accessToken);
      if (claims) {
        const info = {
          name: claims.name || claims.preferred_username || 'Unknown',
          email: claims.preferred_username || claims.upn || claims.email,
          tenantId: claims.tid,
          objectId: claims.oid,
          roles: claims.roles || [],
          // Note: License info is NOT in the token - requires Graph API call
        };
        setUserInfo(info);
        // Cache the decoded info for persistence
        localStorage.setItem('SSO_USER_INFO', JSON.stringify(info));
      }
    }
  }, [accessToken]);

  if (!userInfo) {
    return (
      <div className="p-4 bg-gray-50 rounded-md">
        <p className="text-sm text-gray-600">Loading user information...</p>
      </div>
    );
  }

  return (
    <div className="p-4 rounded-md border border-gray-200 bg-white">
      <h3 className="text-sm font-semibold text-gray-700 mb-3">Account Information</h3>

      <div className="space-y-2 text-sm">
        <div>
          <span className="text-gray-500">Name:</span>{' '}
          <span className="font-medium text-gray-900">{userInfo.name}</span>
        </div>

        {userInfo.email && (
          <div>
            <span className="text-gray-500">Email:</span>{' '}
            <span className="font-medium text-gray-900">{userInfo.email}</span>
          </div>
        )}

        {userInfo.tenantId && (
          <div>
            <span className="text-gray-500">Organization ID:</span>{' '}
            <span className="font-mono text-xs text-gray-700">{userInfo.tenantId}</span>
          </div>
        )}
      </div>

      {/* Note about license information */}
      <div className="mt-4 p-3 bg-blue-50 border border-blue-200 rounded-md">
        <p className="text-xs text-blue-800">
          <strong>Note:</strong> License information requires Microsoft Graph API access.
          Contact your administrator to enable this feature.
        </p>
      </div>
    </div>
  );
}
