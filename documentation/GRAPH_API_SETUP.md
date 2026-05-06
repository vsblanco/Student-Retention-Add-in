# Microsoft Graph API Access for Office Add-ins

## The "Invalid Audience" Problem

Office SSO tokens (`Office.auth.getAccessToken()`) are scoped for your add-in:
- **Audience**: `api://71f37f39-a330-413a-be61-0baa5ce03ea3`

But Microsoft Graph API requires tokens with:
- **Audience**: `https://graph.microsoft.com`

You **cannot** use Office SSO tokens directly to call Graph API. You'll get:
```json
{
  "error": {
    "code": "InvalidAuthenticationToken",
    "message": "Access token validation failure. Invalid audience."
  }
}
```

## Solutions

### Option 1: On-Behalf-Of (OBO) Flow [Recommended by Microsoft]

This is the official Microsoft approach for Office Add-ins accessing Graph API.

#### Architecture
```
Office Add-in (Client)
  ↓ Get Office SSO token
  ↓
Backend Service (Your API)
  ↓ Exchange token using OBO flow
  ↓
Microsoft Graph API
  ↓ Return data
  ↓
Backend Service
  ↓ Return to client
  ↓
Office Add-in
```

#### Implementation Steps

**1. Create a Backend Service**

Options:
- Azure Function (serverless, easy to deploy)
- ASP.NET Web API
- Node.js Express server
- Any backend that can call Microsoft Graph

**2. Configure Azure AD**

Your existing app (`71f37f39-a330-413a-be61-0baa5ce03ea3`) needs:

- **API Permissions**:
  - `User.Read` (Microsoft Graph, Delegated)
  - `User.ReadBasic.All` (for license info)

- **Client Secret or Certificate**:
  - Go to "Certificates & secrets"
  - Create a new client secret
  - Save the secret value (shown only once!)

**3. Backend Code Example (Azure Function - Node.js)**

```javascript
const msal = require('@azure/msal-node');

module.exports = async function (context, req) {
  const userToken = req.headers.authorization?.replace('Bearer ', '');

  if (!userToken) {
    context.res = { status: 401, body: 'No token provided' };
    return;
  }

  const config = {
    auth: {
      clientId: '71f37f39-a330-413a-be61-0baa5ce03ea3',
      clientSecret: process.env.AZURE_CLIENT_SECRET,
      authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID'
    }
  };

  const cca = new msal.ConfidentialClientApplication(config);

  try {
    // Exchange Office SSO token for Graph API token using OBO
    const oboRequest = {
      oboAssertion: userToken,
      scopes: ['https://graph.microsoft.com/User.Read']
    };

    const response = await cca.acquireTokenOnBehalfOf(oboRequest);
    const graphToken = response.accessToken;

    // Call Microsoft Graph API
    const graphResponse = await fetch('https://graph.microsoft.com/v1.0/me/licenseDetails', {
      headers: {
        'Authorization': `Bearer ${graphToken}`
      }
    });

    const licenses = await graphResponse.json();

    context.res = {
      status: 200,
      body: licenses
    };
  } catch (error) {
    context.res = {
      status: 500,
      body: { error: error.message }
    };
  }
};
```

**4. Update Your Add-in**

Modify `licenseChecker.js`:

```javascript
export async function getUserLicenses(accessToken) {
  try {
    // Call YOUR backend, not Graph API directly
    const response = await fetch('https://your-backend.azurewebsites.net/api/GetLicenses', {
      headers: {
        'Authorization': `Bearer ${accessToken}`, // Office SSO token
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`Backend error: ${response.status}`);
    }

    const data = await response.json();
    return data.value || [];
  } catch (error) {
    console.error('Error fetching user licenses:', error);
    throw error;
  }
}
```

---

### Option 2: MSAL.js Direct Authentication [Simpler, No Backend]

Use MSAL.js library to authenticate directly with Graph API (separate from Office SSO).

#### Pros
- No backend needed
- Works entirely client-side
- Simpler to implement

#### Cons
- User authenticates twice (once for Office, once for Graph)
- Less seamless UX

#### Implementation

**1. Install MSAL**

```bash
npm install @azure/msal-browser
```

**2. Configure MSAL**

```javascript
import { PublicClientApplication } from '@azure/msal-browser';

const msalConfig = {
  auth: {
    clientId: '71f37f39-a330-413a-be61-0baa5ce03ea3',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: 'https://vsblanco.github.io/Student-Retention-Add-in/react/dist/index.html'
  }
};

const msalInstance = new PublicClientApplication(msalConfig);
```

**3. Authenticate for Graph API**

```javascript
async function getGraphToken() {
  const loginRequest = {
    scopes: ['User.Read', 'User.ReadBasic.All']
  };

  try {
    // Try silent first
    const response = await msalInstance.acquireTokenSilent(loginRequest);
    return response.accessToken;
  } catch (error) {
    // If silent fails, show popup
    const response = await msalInstance.acquireTokenPopup(loginRequest);
    return response.accessToken;
  }
}
```

**4. Call Graph API**

```javascript
const graphToken = await getGraphToken();
const response = await fetch('https://graph.microsoft.com/v1.0/me/licenseDetails', {
  headers: {
    'Authorization': `Bearer ${graphToken}`
  }
});
```

---

### Option 3: Decode Office SSO Token [Current Approach]

The Office SSO token already contains user information as claims:
- Name
- Email
- Tenant ID
- Object ID

But **NOT** license information (that requires Graph API).

#### What You Can Get

```javascript
// Decode the JWT token
const claims = JSON.parse(atob(token.split('.')[1]));

// Available in token:
claims.name             // User's display name
claims.preferred_username  // User's email
claims.tid              // Tenant ID
claims.oid              // Object ID
claims.upn              // User Principal Name

// NOT available in token:
// - License information
// - Group memberships
// - Extended user properties
```

This is what `UserInfoDisplay.jsx` currently does - no API calls needed!

---

## Recommendation

**For Production**: Implement Option 1 (OBO Flow with backend)
- Most secure
- Best UX (single sign-on)
- Officially recommended by Microsoft

**For Quick Testing**: Use Option 3 (Token Claims)
- Works immediately
- No additional setup
- Limited to info in token

**For Simple Scenarios**: Consider Option 2 (MSAL.js)
- No backend needed
- Full Graph API access
- Acceptable UX for some use cases

---

## References

- [Microsoft Docs: On-Behalf-Of Flow](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
- [Office Add-ins: Authorize to Microsoft Graph](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/authorize-to-microsoft-graph)
- [MSAL.js Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)

---

**Current Status**: The add-in uses Option 3 (token decoding) to show user information without Graph API calls. To enable license checking, implement Option 1 or 2 above.
