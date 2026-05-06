// Example: How to integrate license checking with SSO

import React, { useState } from 'react';
import SSO, { useOfficeSSO } from './SSO';
import LicenseChecker from './LicenseChecker';

export default function SSOWithLicenseCheck({ onNameSelect }) {
  const [accessToken, setAccessToken] = useState(null);
  const [userName, setUserName] = useState(null);

  const handleLoginSuccess = (name, token) => {
    setUserName(name);
    setAccessToken(token);
    if (onNameSelect) {
      onNameSelect(name);
    }
  };

  return (
    <div>
      {!userName ? (
        <SSO onNameSelect={handleLoginSuccess} />
      ) : (
        <div className="p-4">
          <h2 className="text-xl font-bold mb-4">Welcome, {userName}!</h2>

          {/* Show license information */}
          <div className="mb-6">
            <h3 className="text-sm font-semibold text-gray-700 mb-2">License Information</h3>
            <LicenseChecker accessToken={accessToken} />
          </div>

          {/* Rest of your app */}
        </div>
      )}
    </div>
  );
}
