/**
 * SSO Configuration
 *
 * This file controls the Single Sign-On (SSO) behavior for the Office Add-in.
 *
 * Configuration Options:
 * - ENABLE_SSO_FALLBACK: When true, automatically falls back to hardcoded users if SSO fails
 * - FORCE_FALLBACK_MODE: When true, always uses hardcoded users (useful for development/testing)
 * - SSO_RETRY_ATTEMPTS: Number of times to retry SSO before falling back
 * - SHOW_SSO_OPTION: Whether to show the Microsoft SSO button in fallback mode
 */

export const SSOConfig = {
  // If true, tries Microsoft SSO first, then falls back to hardcoded users on failure
  ENABLE_SSO_FALLBACK: true,

  // If true, always uses fallback mode (hardcoded users) - useful for development
  // Set to false to enable Microsoft SSO attempts
  FORCE_FALLBACK_MODE: false,

  // Number of SSO retry attempts before falling back
  SSO_RETRY_ATTEMPTS: 2,

  // Show Microsoft SSO button even in fallback mode (allows manual retry)
  SHOW_SSO_OPTION: true,

  // Timeout for SSO operation (milliseconds)
  SSO_TIMEOUT: 10000,
};

/**
 * Check if SSO is properly configured in the manifest
 * @returns {boolean} True if SSO appears to be configured
 */
export function isSSOConfigured() {
  // Check if we're running in Office environment
  if (typeof Office === 'undefined' || !Office.auth || !Office.auth.getAccessToken) {
    return false;
  }

  // In production, you might want to add more checks here
  // For example, checking if the manifest has a valid WebApplicationInfo section
  return true;
}

/**
 * Determine if we should attempt SSO based on configuration
 * @returns {boolean} True if SSO should be attempted
 */
export function shouldAttemptSSO() {
  if (SSOConfig.FORCE_FALLBACK_MODE) {
    return false;
  }

  return isSSOConfigured();
}
