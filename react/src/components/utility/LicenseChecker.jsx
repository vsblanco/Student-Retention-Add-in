import React, { useState, useEffect } from 'react';
import { checkPowerAutomatePremium } from '../../services/licenseChecker';

export default function LicenseChecker({ accessToken }) {
  const [licenseInfo, setLicenseInfo] = useState(null);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    if (!accessToken) {
      setIsLoading(false);
      return;
    }

    const fetchLicenseInfo = async () => {
      try {
        setIsLoading(true);
        const info = await checkPowerAutomatePremium(accessToken);
        setLicenseInfo(info);
        setError(null);
      } catch (err) {
        console.error('License check failed:', err);
        setError(err.message);
      } finally {
        setIsLoading(false);
      }
    };

    fetchLicenseInfo();
  }, [accessToken]);

  if (!accessToken) {
    return (
      <div className="p-4 bg-gray-50 rounded-md">
        <p className="text-sm text-gray-600">
          Sign in with Microsoft SSO to check your Power Automate license
        </p>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className="p-4 bg-gray-50 rounded-md flex items-center gap-2">
        <div className="animate-spin rounded-full h-4 w-4 border-t-2 border-b-2 border-blue-600"></div>
        <p className="text-sm text-gray-600">Checking license information...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-4 bg-yellow-50 border border-yellow-200 rounded-md">
        <p className="text-sm text-yellow-800">
          ⚠️ Unable to check license: {error}
        </p>
        <p className="text-xs text-yellow-600 mt-1">
          You may need to grant additional permissions in Azure AD
        </p>
      </div>
    );
  }

  if (!licenseInfo) {
    return null;
  }

  return (
    <div className="p-4 rounded-md border">
      {licenseInfo.hasPremium ? (
        <div className="bg-green-50 border-green-200">
          <div className="flex items-center gap-2 mb-2">
            <svg className="w-5 h-5 text-green-600" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
            </svg>
            <span className="font-semibold text-green-800">Power Automate Premium License</span>
          </div>
          <p className="text-sm text-green-700">
            You have access to premium Power Automate features including premium connectors and RPA.
          </p>
        </div>
      ) : licenseInfo.hasPowerAutomate ? (
        <div className="bg-blue-50 border-blue-200">
          <div className="flex items-center gap-2 mb-2">
            <svg className="w-5 h-5 text-blue-600" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clipRule="evenodd" />
            </svg>
            <span className="font-semibold text-blue-800">Power Automate Standard License</span>
          </div>
          <p className="text-sm text-blue-700">
            You have Power Automate included in your Microsoft 365 license. Premium features require an upgrade.
          </p>
        </div>
      ) : (
        <div className="bg-gray-50 border-gray-200">
          <div className="flex items-center gap-2 mb-2">
            <svg className="w-5 h-5 text-gray-600" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
            </svg>
            <span className="font-semibold text-gray-800">No Power Automate License</span>
          </div>
          <p className="text-sm text-gray-700">
            You don't have a Power Automate license. Contact your admin to get access.
          </p>
        </div>
      )}

      {/* Show detected licenses (debugging) */}
      {licenseInfo.licenses && licenseInfo.licenses.length > 0 && (
        <details className="mt-3">
          <summary className="text-xs text-gray-600 cursor-pointer hover:text-gray-800">
            View detected licenses ({licenseInfo.licenses.length})
          </summary>
          <ul className="mt-2 space-y-1">
            {licenseInfo.licenses.map((license, idx) => (
              <li key={idx} className="text-xs text-gray-600 font-mono">
                {license.name}
              </li>
            ))}
          </ul>
        </details>
      )}
    </div>
  );
}
