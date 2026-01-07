/**
 * Azure Function: Exchange Office SSO Token for Microsoft Graph API Token
 * Uses OAuth2 On-Behalf-Of (OBO) flow
 */

module.exports = async function (context, req) {
    // Handle CORS preflight
    if (req.method === 'OPTIONS') {
        context.res = {
            status: 204,
            headers: {
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Methods': 'POST, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type, Authorization',
                'Access-Control-Max-Age': '86400'
            }
        };
        return;
    }

    context.log('Token exchange request received');

    try {
        // Get Office SSO token from request body
        const { token: officeSsoToken } = req.body;

        if (!officeSsoToken) {
            context.res = {
                status: 400,
                headers: {
                    'Access-Control-Allow-Origin': '*',
                    'Content-Type': 'application/json'
                },
                body: { error: 'Missing token in request body' }
            };
            return;
        }

        // Get Azure AD configuration from environment variables
        const tenantId = process.env.AZURE_TENANT_ID;
        const clientId = process.env.AZURE_CLIENT_ID;
        const clientSecret = process.env.AZURE_CLIENT_SECRET;

        if (!tenantId || !clientId || !clientSecret) {
            context.log.error('Missing Azure AD configuration');
            context.res = {
                status: 500,
                headers: {
                    'Access-Control-Allow-Origin': '*',
                    'Content-Type': 'application/json'
                },
                body: { error: 'Server configuration error' }
            };
            return;
        }

        // Exchange token using OAuth2 On-Behalf-Of flow
        const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

        const params = new URLSearchParams({
            grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
            client_id: clientId,
            client_secret: clientSecret,
            assertion: officeSsoToken,
            scope: 'https://graph.microsoft.com/Mail.Send',
            requested_token_use: 'on_behalf_of'
        });

        context.log('Exchanging token with Azure AD...');

        const response = await fetch(tokenEndpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: params.toString()
        });

        const data = await response.json();

        if (!response.ok) {
            context.log.error('Token exchange failed:', data);
            context.res = {
                status: response.status,
                headers: {
                    'Access-Control-Allow-Origin': '*',
                    'Content-Type': 'application/json'
                },
                body: {
                    error: 'Token exchange failed',
                    details: data.error_description || data.error
                }
            };
            return;
        }

        context.log('Token exchange successful');

        // Return the Graph API access token
        context.res = {
            status: 200,
            headers: {
                'Access-Control-Allow-Origin': '*',
                'Content-Type': 'application/json'
            },
            body: {
                accessToken: data.access_token,
                expiresIn: data.expires_in,
                tokenType: data.token_type
            }
        };

    } catch (error) {
        context.log.error('Unexpected error:', error);
        context.res = {
            status: 500,
            headers: {
                'Access-Control-Allow-Origin': '*',
                'Content-Type': 'application/json'
            },
            body: {
                error: 'Internal server error',
                message: error.message
            }
        };
    }
};
