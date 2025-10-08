import React, { useState } from "react";

export function useOfficeSSO() {
  const [token, setToken] = useState(null);
  const [error, setError] = useState(null);

  async function getAccessToken() {
    try {
      const accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
      setToken(accessToken);
      setError(null);
      return accessToken;
    } catch (err) {
      setError(err.message);
      setToken(null);
      return null;
    }
  }

  return { token, error, getAccessToken };
}

export default function SSO({ onNameSelect }) {
  const [showName, setShowName] = useState(false);
  const [loginStatus, setLoginStatus] = useState("");
  const { getAccessToken, token, error } = useOfficeSSO();

  const handleTestClick = () => {
    setShowName(true);
    if (onNameSelect) {
      onNameSelect("Victor Blanco");
    }
  };

  const handleSSOLogin = async () => {
    const accessToken = await getAccessToken();
    if (accessToken) {
      setLoginStatus(`Logged in as ${accessToken}`);
      if (onNameSelect) {
        onNameSelect(accessToken);
      }
    } else {
      setLoginStatus(`Login failed${error ? `: ${error}` : ""}`);
    }
  };

  return (
    <div className="flex flex-col items-center">
      <div className="flex space-x-2">
        <button
          className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 transition"
          onClick={handleTestClick}
        >
          Test
        </button>
        <button
          className="px-4 py-2 bg-gray-800 text-white rounded hover:bg-gray-900 transition flex items-center space-x-2"
          onClick={handleSSOLogin}
        >
          <span>
            <svg width="20" height="20" viewBox="0 0 20 20" className="mr-2" xmlns="http://www.w3.org/2000/svg">
              <rect x="1" y="1" width="8" height="8" fill="#F25022"/>
              <rect x="11" y="1" width="8" height="8" fill="#7FBA00"/>
              <rect x="1" y="11" width="8" height="8" fill="#00A4EF"/>
              <rect x="11" y="11" width="8" height="8" fill="#FFB900"/>
            </svg>
          </span>
          <span>Microsoft SSO</span>
        </button>
      </div>
      {showName && <div className="mt-4">Victor Blanco</div>}
      {loginStatus && <div className="mt-4">{loginStatus}</div>}
    </div>
  );
}
