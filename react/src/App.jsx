// Timestamp: 2025-12-15 12:05:00 | Version: 2.2.0
import React, { useState, useEffect, useRef, lazy, Suspense } from 'react';
import StudentView from './components/studentView/StudentView.jsx';
import About from './components/about/About.jsx';
import Settings from './components/settings/Settings.jsx';
import LDAManager from './components/createLDA/LDAManager.jsx';
import PersonalizedEmail from './components/personalizedEmail/PersonalizedEmail.jsx';
import ReportGeneration from './components/reportGeneration/ReportGeneration.jsx';
import Welcome from './components/welcomeScreen/Welcome.jsx'; // Import the Welcome component
import chromeExtensionService from '../../shared/chromeExtensionService.js';
import { registerWorkbookUser, isUserRegistered } from './services/workbookUsers.js';
import { ToastContainer, Slide } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

// Lazy load SSO to ensure it doesn't block initial render
const SSO = lazy(() => import('./components/utility/SSO.jsx'));

function App() {
  const [page, setPage] = useState(null);
  const [isLoading, setIsLoading] = useState(true);
  const [currentUser, setCurrentUser] = useState(null);
  const [accessToken, setAccessToken] = useState(null); // <-- ADDED: Store SSO access token

  // Tutorial/Welcome State
  const [showWelcome, setShowWelcome] = useState(false);

  // Ref to track if we've already handled the "Ready" signal for the current page load
  const isReadyRef = useRef(false);

  // --- 1. AUTH CHECK ON MOUNT ---
  useEffect(() => {
    const checkAuth = () => {
      try {
        // Standardized to use 'SSO_USER'
        const cachedUser = localStorage.getItem('SSO_USER');
        
        if (cachedUser) {
          console.log("App: Auto-logged in as", cachedUser);
          setCurrentUser(cachedUser);
          // We keep isLoading = true here because we are about to mount StudentView,
          // which will signal when it's ready via onReady.
        } else {
          console.log("App: No session found. Showing Login.");
          // Stop loading so the SSO screen is visible
          setIsLoading(false);
        }
      } catch (e) {
        console.warn("Auth check failed", e);
        setIsLoading(false);
      }
    };

    checkAuth();
  }, []);

  // --- 2. PAGE ROUTING & TUTORIAL CHECK ---
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const pageParam = params.get('page');
    const targetPage = pageParam || 'student-view';
    setPage(targetPage);

    // Check if user has seen welcome screen (Simulating "New User" detection)
    const hasSeenWelcome = localStorage.getItem('SRK_HAS_SEEN_WELCOME');
    if (!hasSeenWelcome) {
      setShowWelcome(true);
    }

    // Safety Timeout: If StudentView (or any feature) hangs, remove spinner after 8s
    // Only applies if we are actually logged in and trying to load a feature
    if (currentUser) {
        const safetyTimer = setTimeout(() => {
            if (!isReadyRef.current && isLoading) {
                console.warn("Feature took too long. Forcing load.");
                setIsLoading(false);
            }
        }, 8000);
        return () => clearTimeout(safetyTimer);
    }
  }, [currentUser, isLoading]); 

  // --- HANDLERS ---

  // Called by Child Components (StudentView) when Excel data is bound
  const handleFeatureReady = () => {
    if (!isReadyRef.current) {
      isReadyRef.current = true;
      setIsLoading(false);
    }
  };

  // Called by SSO component on successful login
  const handleLoginSuccess = (username, token = null) => { // <-- ADDED: Accept token parameter
    console.log("App: Login Successful ->", username);
    setCurrentUser(username);
    setAccessToken(token); // <-- ADDED: Store access token

    // Set loading to TRUE to prevent UI flicker while StudentView mounts
    setIsLoading(true);
    isReadyRef.current = false;

    // Standardized save to 'SSO_USER'
    localStorage.setItem('SSO_USER', username);
  };

  const handleWelcomeClose = () => {
    setShowWelcome(false);
    // Mark as seen so it doesn't appear again
    localStorage.setItem('SRK_HAS_SEEN_WELCOME', 'true');
  };

  // Detect whether the signed-in user is new to this workbook and register
  // them if so. Runs on Office.onReady (covers fresh loads + cached auto-login),
  // re-runs on SettingsChanged (co-author updates the list) and on
  // visibilitychange (covers Excel Online cases where the task pane persists
  // across workbook switches without remounting React). Pulls email from the
  // SSO_USER_INFO cache SSO.jsx writes during login; falls back to the bare
  // name. registerWorkbookUser is idempotent, so repeated calls are safe.
  useEffect(() => {
    if (!currentUser) return;
    if (typeof window === 'undefined' || !window.Office || !Office.onReady) return;

    const resolveUserInfo = () => {
      let userInfo = { name: currentUser };
      try {
        const cached = localStorage.getItem('SSO_USER_INFO');
        if (cached) {
          const parsed = JSON.parse(cached);
          if (parsed && typeof parsed === 'object') {
            userInfo = {
              name: parsed.name || currentUser,
              email: typeof parsed.email === 'string' ? parsed.email : undefined,
            };
          }
        }
      } catch (e) {
        console.warn('App: failed to parse SSO_USER_INFO for workbook user registration', e);
      }
      return userInfo;
    };

    const detectAndRegister = () => {
      const userInfo = resolveUserInfo();
      let isNew = false;
      try {
        isNew = !isUserRegistered(userInfo);
      } catch (e) {
        // If the membership check fails (e.g. settings unavailable), fall
        // through to registerWorkbookUser — it's idempotent and will no-op
        // if the user is already present.
        console.warn('App: workbook user membership check failed', e);
      }
      console.log(
        `App: ${isNew ? 'new' : 'returning'} user "${userInfo.name}" — ${isNew ? 'registering in' : 'already in'} workbook`
      );
      registerWorkbookUser(userInfo).catch(err => {
        console.warn('App: failed to register workbook user', err);
      });
    };

    let unmounted = false;
    let settingsHandler = null;
    let visibilityHandler = null;

    Office.onReady(() => {
      if (unmounted) return;

      // Initial check on app load (or login).
      detectAndRegister();

      // SettingsChanged only fires when a co-author saves settings in Excel
      // on the web, so refresh the in-memory copy before re-checking.
      try {
        const settings = Office.context && Office.context.document && Office.context.document.settings;
        if (settings && typeof settings.addHandlerAsync === 'function' &&
            Office.EventType && Office.EventType.SettingsChanged) {
          settingsHandler = () => {
            if (typeof settings.refreshAsync === 'function') {
              settings.refreshAsync(() => { if (!unmounted) detectAndRegister(); });
            } else {
              detectAndRegister();
            }
          };
          settings.addHandlerAsync(Office.EventType.SettingsChanged, settingsHandler);
        }
      } catch (e) {
        console.warn('App: failed to attach SettingsChanged listener', e);
      }

      // Visibility change covers the Excel Online workbook-switch case where
      // the task pane persists and Office.onReady doesn't re-fire.
      if (typeof document !== 'undefined' && document.addEventListener) {
        visibilityHandler = () => {
          if (document.visibilityState === 'visible') detectAndRegister();
        };
        document.addEventListener('visibilitychange', visibilityHandler);
      }
    });

    return () => {
      unmounted = true;
      if (settingsHandler) {
        try {
          Office.context.document.settings.removeHandlerAsync(
            Office.EventType.SettingsChanged,
            settingsHandler
          );
        } catch { /* ignore */ }
      }
      if (visibilityHandler && typeof document !== 'undefined') {
        document.removeEventListener('visibilitychange', visibilityHandler);
      }
    };
  }, [currentUser]);

  // --- CHROME EXTENSION MASTER RELAY ---
  // App.jsx manages the Chrome extension connection for all child components
  useEffect(() => {
    console.log("App: Initializing Chrome Extension Service (Master Relay)");

    // Start extension detection
    chromeExtensionService.startPinging();

    // Start keep-alive heartbeat
    chromeExtensionService.startKeepAlive();

    // Listen for extension events
    const removeListener = chromeExtensionService.addListener((event) => {
      if (event.type === "installed") {
        console.log("App: Chrome Extension is installed and ready!");
      }
    });

    // Cleanup on unmount
    return () => {
      console.log("App: Cleaning up Chrome Extension Service");
      removeListener();
      chromeExtensionService.stopPinging();
      chromeExtensionService.stopKeepAlive();
    };
  }, []);

  // --- RENDER HELPERS ---
  const renderContent = () => {
    // A. If no user, show SSO (Login)
    if (!currentUser) {
      return (
        <Suspense fallback={<div className="p-10 text-center">Loading Login...</div>}>
          <SSO onNameSelect={handleLoginSuccess} />
        </Suspense>
      );
    }

    // B. If user exists, show the requested page
    switch (page) {
      case 'settings':
        return <Settings user={currentUser} accessToken={accessToken} onReady={handleFeatureReady} />;
case 'about':
        return <About onReady={handleFeatureReady} />;
      case 'create-lda':
        return <LDAManager user={currentUser} onReady={handleFeatureReady} />;
      case 'personalized-email':
        return <PersonalizedEmail user={currentUser} accessToken={accessToken} onReady={handleFeatureReady} />;
      case 'report-generation':
        return <ReportGeneration user={currentUser} onReady={handleFeatureReady} />;
      case 'student-view':
      default:
        // Pass user and ready-handler to StudentView
        return <StudentView user={currentUser} onReady={handleFeatureReady} />;
    }
  };

  return (
    <>
      {/* LOADING OVERLAY */}
      <div 
        style={{
          position: 'fixed',
          inset: 0,
          zIndex: 9999,
          backgroundColor: '#f9fafb',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          transition: 'opacity 0.7s ease-out',
          opacity: isLoading ? 1 : 0,
          pointerEvents: isLoading ? 'auto' : 'none',
        }}
      >
        <div className="flex flex-col items-center gap-4">
           <div className="animate-spin rounded-full h-12 w-12 border-t-4 border-b-4 border-blue-600"></div>
           <span className="text-gray-500 font-medium text-sm animate-pulse">
             {currentUser ? "Loading your dashboard..." : "Preparing..."}
           </span>
        </div>
      </div>

      {/* MAIN APP CONTENT (Login or Feature) */}
      {renderContent()}

      {/* TUTORIAL / WELCOME OVERLAY 
          Only shows if:
          1. App is done loading (!isLoading)
          2. User is logged in (currentUser exists)
          3. User hasn't seen it yet (showWelcome is true)
      */}
      {!isLoading && currentUser && showWelcome && (
        <Welcome onClose={handleWelcomeClose} user={currentUser} />
      )}
      
      <ToastContainer
        position="bottom-left"
        pauseOnFocusLoss={false}
        transition={Slide}
        limit={3}
      />
    </>
  );
}

export default App;