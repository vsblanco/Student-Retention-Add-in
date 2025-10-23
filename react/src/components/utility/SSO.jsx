// 2025-10-08T23:40:56.531Z - v1.0.0
import React, { useState } from "react";
import { ToastContainer, toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import SSOtemp from "./SSOtemp"; // <-- added import

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
    return decoded.name || "Name not found in token";
  } catch (e) {
    console.error("Failed to decode JWT:", e);
    return "Error decoding token";
  }
}

export function useOfficeSSO() {
  const [token, setToken] = useState(null);
  const [error, setError] = useState(null);

  async function getAccessToken() {
    try {
      // ⚠️ CORRECTED: Use Office.auth for task pane SSO, not OfficeRuntime.auth
      if (!window.Office || !Office.auth || !Office.auth.getAccessToken) {
        throw new Error("The identity API is not supported for this add-in.");
      }
      const accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
      setToken(accessToken);
      setError(null);
      return accessToken;
    } catch (err) {
      // In case of error, err.code is often more useful than err.message
      setError(`Error code ${err.code}: ${err.message}`);
      setToken(null);
      return null;
    }
  }

  return { token, error, getAccessToken };
}

const CACHE_KEY = "SSO_USER"; // <-- new cache key

const test = true;

export default function SSO({ onNameSelect }) {
  const { getAccessToken, error } = useOfficeSSO();

  const handleTestClick = () => {
    const userName = "Victor Blanco";
    toast.success(`Welcome back ${userName}`, { position: "bottom-center" });
    if (onNameSelect) {
      onNameSelect(userName);
    }
  };

  const handleSSOLogin = async () => {
    const accessToken = await getAccessToken();
    if (accessToken) {
      // ✅ CORRECTED: Decode the token to get the user's name
      const userName = decodeJwt(accessToken);
      // Cache the SSO user
      try {
        if (typeof window !== "undefined" && window.localStorage) {
          localStorage.setItem(CACHE_KEY, userName);
        }
      } catch (e) {
        console.warn("Failed to write sso user to localStorage", e);
      }
      toast.success(`Success! Logged in as: ${userName}`, { position: "bottom-center" });
      if (onNameSelect) {
        onNameSelect(userName);
      }
    } else {
      toast.error(`Login failed${error ? `: ${error}` : ""}`, { position: "bottom-center" });
    }
  };

  // New handler to receive selection from SSOtemp
  const handleTempSelect = (userName) => {
    // Cache the selected user
    try {
      if (typeof window !== "undefined" && window.localStorage) {
        localStorage.setItem(CACHE_KEY, userName);
      }
    } catch (e) {
      console.warn("Failed to write sso user to localStorage", e);
    }

    toast.success(`Welcome back ${userName}`, { position: "bottom-center" });
    if (onNameSelect) {
      onNameSelect(userName);
    }
  };

  // If test mode, render the SSOtemp UI to pick a user
  if (test) {
    // pass cached user as defaultUser so SSOtemp can initialize from it
    const cachedUser = (typeof window !== "undefined" && window.localStorage) ? localStorage.getItem(CACHE_KEY) : null;
    return (
      <div className="flex flex-col items-center justify-center min-h-[60vh]">
        <SSOtemp onSelect={handleTempSelect} defaultUser={cachedUser} />
        <ToastContainer />
      </div>
    );
  }

  return (
    <div className="flex flex-col items-center justify-center min-h-[60vh]">
      <div className="mb-8 text-center">
        <h1 className="text-2xl font-bold mb-2">Welcome!</h1>
        <p className="text-lg text-gray-600">Please sign in to get started.</p>
      </div>
      {/* Microsoft SSO button */}
      <div className="flex flex-col items-center mb-4 space-y-2 w-64">
        <button
          className="w-full px-4 py-2 bg-gray-800 text-white rounded hover:bg-gray-900 transition flex items-center justify-center space-x-2"
          onClick={handleSSOLogin}
        >
          <span>
            <svg width="20" height="20" viewBox="0 0 20 20" className="mr-2" xmlns="http://www.w3.org/2000/svg">
              <rect x="1" y="1" width="8" height="8" fill="#F25022" />
              <rect x="11" y="1" width="8" height="8" fill="#7FBA00" />
              <rect x="1" y="11" width="8" height="8" fill="#00A4EF" />
              <rect x="11" y="11" width="8" height="8" fill="#FFB900" />
            </svg>
          </span>
          <span>Microsoft SSO</span>
        </button>
        {/* Test button below */}
        <button
          className="w-full px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 transition font-medium"
          onClick={handleTestClick}
        >
          Debug User
        </button>
      </div>
      <ToastContainer />
    </div>
  );
}