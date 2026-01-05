// 2025-12-26 - v2.1.0 - Microsoft SSO with Guest Fallback
import React, { useState, useEffect } from "react";
import { ToastContainer, toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import { SSOConfig, shouldAttemptSSO } from "../../config/ssoConfig";

// Helper function to decode the JWT token and get user name
function decodeJwt(token) {
  try {
    const base64Url = token.split(".")[1];
    const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
    const jsonPayload = decodeURIComponent(
      atob(base64)
        .split("")
        .map(function (c) {
          return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
        })
        .join("")
    );

    const decoded = JSON.parse(jsonPayload);
    // The 'name' claim contains the user's full name
    return decoded.name || decoded.preferred_username || "Unknown User";
  } catch (e) {
    console.error("Failed to decode JWT:", e);
    return null;
  }
}

// Helper function to decode and cache full user info from token
function cacheUserInfoFromToken(token) {
  try {
    const base64Url = token.split(".")[1];
    const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
    const jsonPayload = decodeURIComponent(
      atob(base64)
        .split("")
        .map(function (c) {
          return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
        })
        .join("")
    );

    const claims = JSON.parse(jsonPayload);

    // Extract and cache user information
    const userInfo = {
      name: claims.name || claims.preferred_username || 'Unknown',
      email: claims.preferred_username || claims.upn || claims.email,
      tenantId: claims.tid,
      objectId: claims.oid,
      roles: claims.roles || [],
    };

    // Cache to localStorage for persistence across features
    if (typeof window !== "undefined" && window.localStorage) {
      localStorage.setItem('SSO_USER_INFO', JSON.stringify(userInfo));
      console.log('SSO: User info cached to localStorage');
    }

    return userInfo;
  } catch (e) {
    console.error("Failed to cache user info from token:", e);
    return null;
  }
}

// Helper function to save user info to Excel workbook settings
async function saveUserToWorkbookSettings(userInfo) {
  if (!userInfo || !userInfo.email) {
    console.warn('SSO: Cannot save user - invalid user info');
    return;
  }

  try {
    await Excel.run(async (context) => {
      const settings = context.workbook.settings;
      const usersSetting = settings.getItemOrNullObject("sso_users");
      usersSetting.load("value");
      await context.sync();

      let users = [];
      if (usersSetting.value) {
        try {
          users = JSON.parse(usersSetting.value);
          if (!Array.isArray(users)) {
            users = [];
          }
        } catch (e) {
          console.error('SSO: Failed to parse existing users', e);
          users = [];
        }
      }

      // Check if user already exists (by email or objectId)
      const existingUserIndex = users.findIndex(u =>
        u.email === userInfo.email ||
        (userInfo.objectId && u.objectId === userInfo.objectId)
      );

      const userRecord = {
        name: userInfo.name,
        email: userInfo.email,
        tenantId: userInfo.tenantId,
        objectId: userInfo.objectId,
        roles: userInfo.roles || [],
        lastLogin: new Date().toISOString()
      };

      if (existingUserIndex !== -1) {
        // Update existing user record
        users[existingUserIndex] = userRecord;
        console.log('SSO: Updated existing user in workbook settings');
      } else {
        // Add new user record
        users.push(userRecord);
        console.log('SSO: Added new user to workbook settings');
      }

      // Save back to settings
      settings.add("sso_users", JSON.stringify(users));
      await context.sync();

      console.log('SSO: User info saved to workbook settings');
    });
  } catch (error) {
    console.error('SSO: Failed to save user to workbook settings', error);
    // Don't throw - this shouldn't block login
  }
}

export function useOfficeSSO() {
  const [token, setToken] = useState(null);
  const [error, setError] = useState(null);
  const [isLoading, setIsLoading] = useState(false);

  async function getAccessToken(options = {}) {
    const { silent = false, timeout = SSOConfig.SSO_TIMEOUT } = options;

    setIsLoading(true);
    try {
      // Check if Office.auth is available
      if (!window.Office || !Office.auth || !Office.auth.getAccessToken) {
        throw new Error("Office SSO API is not available");
      }

      // Create a timeout promise
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("SSO timeout")), timeout)
      );

      // Race between SSO and timeout
      const accessToken = await Promise.race([
        Office.auth.getAccessToken({
          allowSignInPrompt: !silent,
          allowConsentPrompt: !silent,
          forMSGraphAccess: true
        }),
        timeoutPromise
      ]);

      setToken(accessToken);
      setError(null);
      setIsLoading(false);
      return accessToken;
    } catch (err) {
      const errorCode = err.code || 0;
      const errorMsg = err.code ? `Error ${err.code}: ${err.message}` : err.message;
      console.error("SSO Error:", errorMsg);

      // Errors 13000-13012 are configuration errors - don't show to user if fallback is enabled
      const isConfigError = errorCode >= 13000 && errorCode <= 13012;
      if (!isConfigError || !SSOConfig.ENABLE_SSO_FALLBACK) {
        setError(errorMsg);
      }

      setToken(null);
      setIsLoading(false);
      return null;
    }
  }

  return { token, error, isLoading, getAccessToken };
}

const CACHE_KEY = "SSO_USER";
const CACHE_MODE_KEY = "SSO_MODE"; // Track which mode was used for login

export default function SSO({ onNameSelect }) {
  const { getAccessToken, error, isLoading } = useOfficeSSO();
  const [mode, setMode] = useState("auto"); // "auto", "sso"
  const [ssoAttempted, setSsoAttempted] = useState(false);
  const [autoLoginAttempted, setAutoLoginAttempted] = useState(false);

  // Auto-attempt SSO on mount (if configured)
  useEffect(() => {
    if (autoLoginAttempted) return;

    const attemptAutoSSO = async () => {
      setAutoLoginAttempted(true);

      // Check if we should attempt SSO
      if (!shouldAttemptSSO()) {
        console.log("SSO: Force fallback mode enabled or SSO not configured");
        setMode("sso");
        return;
      }

      console.log("SSO: Auto-attempting silent SSO...");
      try {
        const accessToken = await getAccessToken({ silent: true });
        if (accessToken) {
          const userName = decodeJwt(accessToken);
          if (userName) {
            console.log("SSO: Silent SSO succeeded");
            cacheUser(userName, "sso");
            const userInfo = cacheUserInfoFromToken(accessToken); // Cache full user info

            // Save user info to workbook settings
            if (userInfo) {
              await saveUserToWorkbookSettings(userInfo);
            }

            toast.success(`Welcome back, ${userName}!`, { position: "bottom-center" });
            if (onNameSelect) {
              onNameSelect(userName, accessToken); // Pass token to parent
            }
            setMode("sso");
            return;
          }
        }
      } catch (err) {
        console.log("SSO: Silent SSO failed, showing login options");
      }

      // If silent SSO fails, show SSO UI
      setMode("sso");
      setSsoAttempted(true);
    };

    attemptAutoSSO();
  }, [autoLoginAttempted]);

  // Helper to cache user
  const cacheUser = (userName, loginMode) => {
    try {
      if (typeof window !== "undefined" && window.localStorage) {
        localStorage.setItem(CACHE_KEY, userName);
        localStorage.setItem(CACHE_MODE_KEY, loginMode);
      }
    } catch (e) {
      console.warn("Failed to cache user to localStorage", e);
    }
  };

  // Handle Microsoft SSO login
  const handleSSOLogin = async () => {
    console.log("SSO: Manual SSO login triggered");
    const accessToken = await getAccessToken({ silent: false });
    if (accessToken) {
      const userName = decodeJwt(accessToken);
      if (userName) {
        cacheUser(userName, "sso");
        const userInfo = cacheUserInfoFromToken(accessToken); // Cache full user info

        // Save user info to workbook settings
        if (userInfo) {
          await saveUserToWorkbookSettings(userInfo);
        }

        toast.success(`Success! Logged in as: ${userName}`, { position: "bottom-center" });
        if (onNameSelect) {
          onNameSelect(userName, accessToken); // Pass token to parent
        }
        setMode("sso");
      } else {
        toast.error("Failed to decode user information from token", { position: "bottom-center" });
      }
    } else {
      toast.error(`SSO login failed${error ? `: ${error}` : ""}`, { position: "bottom-center" });
      setSsoAttempted(true);
    }
  };

  // Handle Guest login
  const handleGuestLogin = () => {
    const guestName = "Guest";
    cacheUser(guestName, "guest");
    toast.success("Continuing as Guest", { position: "bottom-center" });
    if (onNameSelect) {
      onNameSelect(guestName, null); // No access token for guest
    }
  };

  // Render SSO mode
  return (
    <div className="flex flex-col items-center justify-center min-h-[60vh]">
      <div className="mb-8 text-center">
        <h1 className="text-2xl font-bold mb-2">Welcome!</h1>
        <p className="text-lg text-gray-600">Please sign in to get started.</p>
      </div>

      {/* Microsoft SSO button */}
      <div className="flex flex-col items-center mb-4 space-y-2 w-64">
        <button
          className="w-full px-4 py-2 bg-gray-800 text-white rounded hover:bg-gray-900 transition flex items-center justify-center space-x-2 disabled:opacity-50 disabled:cursor-not-allowed"
          onClick={handleSSOLogin}
          disabled={isLoading}
        >
          {isLoading ? (
            <>
              <div className="animate-spin rounded-full h-4 w-4 border-t-2 border-b-2 border-white"></div>
              <span>Signing in...</span>
            </>
          ) : (
            <>
              <svg width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
                <rect x="1" y="1" width="8" height="8" fill="#F25022" />
                <rect x="11" y="1" width="8" height="8" fill="#7FBA00" />
                <rect x="1" y="11" width="8" height="8" fill="#00A4EF" />
                <rect x="11" y="11" width="8" height="8" fill="#FFB900" />
              </svg>
              <span>Sign in with Microsoft</span>
            </>
          )}
        </button>

        {/* Guest login option - always available */}
        <button
          className="w-full px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition font-medium"
          onClick={handleGuestLogin}
        >
          Continue as Guest
        </button>
      </div>

      {/* Show error if any */}
      {error && (
        <div className="mt-4 p-3 bg-red-50 border border-red-200 rounded-md text-red-700 text-sm max-w-md">
          <p className="font-semibold">Authentication Error</p>
          <p className="mt-1">{error}</p>
        </div>
      )}

      <ToastContainer />
    </div>
  );
}