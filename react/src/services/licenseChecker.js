/**
 * Microsoft License Checker
 *
 * Checks if the current user has specific Microsoft licenses
 * using the Microsoft Graph API.
 */

// Power Automate License SKU IDs
// Reference: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
const LICENSE_SKUS = {
  // Power Automate Premium (Per User)
  POWER_AUTOMATE_PREMIUM: 'f30db892-07e9-47e9-837c-80727f46fd3d',

  // Power Automate Premium (Per Flow)
  POWER_AUTOMATE_PER_FLOW: 'a403ebcc-fae0-4ca2-8c8c-7a907fd6c235',

  // Office 365 E3 (includes Power Automate)
  OFFICE_365_E3: '6fd2c87f-b296-42f0-b197-1e91e994b900',

  // Office 365 E5 (includes Power Automate Premium)
  OFFICE_365_E5: 'c7df2760-2c81-4ef7-b578-5b5392b571df',

  // Microsoft 365 E3
  M365_E3: 'SPE_E3',

  // Microsoft 365 E5
  M365_E5: 'SPE_E5',
};

/**
 * Get the user's licenses from Microsoft Graph
 * @param {string} accessToken - The access token from Office SSO
 * @returns {Promise<Array>} Array of license details
 */
export async function getUserLicenses(accessToken) {
  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/me/licenseDetails', {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    return data.value || [];
  } catch (error) {
    console.error('Error fetching user licenses:', error);
    throw error;
  }
}

/**
 * Check if user has Power Automate Premium license
 * @param {string} accessToken - The access token from Office SSO
 * @returns {Promise<Object>} License information
 */
export async function checkPowerAutomatePremium(accessToken) {
  try {
    const licenses = await getUserLicenses(accessToken);

    // Check for Power Automate Premium or E5 licenses
    const hasPremium = licenses.some(license =>
      license.skuId === LICENSE_SKUS.POWER_AUTOMATE_PREMIUM ||
      license.skuId === LICENSE_SKUS.POWER_AUTOMATE_PER_FLOW ||
      license.skuId === LICENSE_SKUS.OFFICE_365_E5 ||
      license.skuPartNumber === LICENSE_SKUS.M365_E5
    );

    // Check for any Power Automate license
    const hasPowerAutomate = licenses.some(license =>
      license.skuId === LICENSE_SKUS.POWER_AUTOMATE_PREMIUM ||
      license.skuId === LICENSE_SKUS.POWER_AUTOMATE_PER_FLOW ||
      license.skuId === LICENSE_SKUS.OFFICE_365_E3 ||
      license.skuId === LICENSE_SKUS.OFFICE_365_E5 ||
      license.skuPartNumber === LICENSE_SKUS.M365_E3 ||
      license.skuPartNumber === LICENSE_SKUS.M365_E5
    );

    return {
      hasPowerAutomate,
      hasPremium,
      licenses: licenses.map(l => ({
        name: l.skuPartNumber,
        id: l.skuId
      }))
    };
  } catch (error) {
    console.error('Error checking Power Automate license:', error);
    return {
      hasPowerAutomate: false,
      hasPremium: false,
      error: error.message,
      licenses: []
    };
  }
}

/**
 * Get user profile with license information
 * @param {string} accessToken - The access token from Office SSO
 * @returns {Promise<Object>} User profile with license info
 */
export async function getUserProfileWithLicenses(accessToken) {
  try {
    // Get user profile
    const profileResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    if (!profileResponse.ok) {
      throw new Error(`Graph API error: ${profileResponse.status}`);
    }

    const profile = await profileResponse.json();

    // Get license information
    const licenseInfo = await checkPowerAutomatePremium(accessToken);

    return {
      name: profile.displayName,
      email: profile.mail || profile.userPrincipalName,
      ...licenseInfo
    };
  } catch (error) {
    console.error('Error getting user profile with licenses:', error);
    throw error;
  }
}
