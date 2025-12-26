import React, { useState, useEffect } from 'react';

/**
 * UserInfoDisplay component
 * Displays cached user information from SSO login
 * Note: User info is cached by SSO.jsx during login
 */
export default function UserInfoDisplay({ accessToken }) {
  const [userInfo, setUserInfo] = useState(null);

  useEffect(() => {
    // Load cached user info from localStorage
    // This is cached by SSO.jsx during login, so it's available across all features
    const cachedInfo = localStorage.getItem('SSO_USER_INFO');
    if (cachedInfo) {
      try {
        setUserInfo(JSON.parse(cachedInfo));
      } catch (e) {
        console.error('Failed to parse cached user info:', e);
      }
    }
  }, []); // Only run once on mount

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
