// 2025-12-26 - v2.0.0 - Microsoft SSO with Intelligent Fallback
import React, { useState, useEffect } from "react";
import { ToastContainer, toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import SSOtemp from "./SSOtemp";
import { SSOConfig, shouldAttemptSSO } from "../../config/ssoConfig";

// Helper function to decode the JWT token
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
  const [mode, setMode] = useState("auto"); // "auto", "sso", "fallback"
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
        setMode("fallback");
        return;
      }

      if (SSOConfig.ENABLE_SSO_FALLBACK) {
        console.log("SSO: Auto-attempting silent SSO...");
        try {
          const accessToken = await getAccessToken({ silent: true });
          if (accessToken) {
            const userName = decodeJwt(accessToken);
            if (userName) {
              console.log("SSO: Silent SSO succeeded");
              cacheUser(userName, "sso");
              toast.success(`Welcome back, ${userName}!`, { position: "bottom-center" });
              if (onNameSelect) {
                onNameSelect(userName, accessToken); // <-- ADDED: Pass token to parent
              }
              setMode("sso");
              return;
            }
          }
        } catch (err) {
          console.log("SSO: Silent SSO failed, falling back to user selection");
        }

        // If silent SSO fails, show fallback
        setMode("fallback");
        setSsoAttempted(true);
      } else {
        // If fallback not enabled, show SSO UI
        setMode("sso");
      }
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
        toast.success(`Success! Logged in as: ${userName}`, { position: "bottom-center" });
        if (onNameSelect) {
          onNameSelect(userName, accessToken); // <-- ADDED: Pass token to parent
        }
        setMode("sso");
      } else {
        toast.error("Failed to decode user information from token", { position: "bottom-center" });
      }
    } else {
      toast.error(`SSO login failed${error ? `: ${error}` : ""}`, { position: "bottom-center" });

      // If SSO fails and fallback is enabled, offer fallback
      if (SSOConfig.ENABLE_SSO_FALLBACK && !ssoAttempted) {
        setSsoAttempted(true);
        toast.info("Switching to fallback login method...", { position: "bottom-center" });
        setTimeout(() => setMode("fallback"), 1500);
      }
    }
  };

  // Handle fallback user selection
  const handleTempSelect = (userName) => {
    cacheUser(userName, "fallback");
    toast.success(`Welcome back, ${userName}`, { position: "bottom-center" });
    if (onNameSelect) {
      onNameSelect(userName);
    }
  };

  // Render fallback mode (hardcoded users)
  if (mode === "fallback") {
    const cachedUser = (typeof window !== "undefined" && window.localStorage)
      ? localStorage.getItem(CACHE_KEY)
      : null;

    return (
      <div className="flex flex-col items-center justify-center min-h-[60vh]">
        <SSOtemp onSelect={handleTempSelect} defaultUser={cachedUser} />

        {/* Show SSO option if enabled */}
        {SSOConfig.SHOW_SSO_OPTION && (
          <div className="mt-4">
            <button
              className="px-4 py-2 text-sm text-gray-600 hover:text-gray-800 underline"
              onClick={() => setMode("sso")}
            >
              Try Microsoft SSO instead
            </button>
          </div>
        )}

        <ToastContainer />
      </div>
    );
  }

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

        {/* Show fallback option if SSO was attempted and failed */}
        {SSOConfig.ENABLE_SSO_FALLBACK && ssoAttempted && (
          <button
            className="w-full px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 transition font-medium"
            onClick={() => setMode("fallback")}
          >
            Use Demo Accounts
          </button>
        )}
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