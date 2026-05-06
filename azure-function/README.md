# Token Exchange Azure Function

This Azure Function exchanges Office SSO tokens for Microsoft Graph API tokens using the OAuth2 On-Behalf-Of (OBO) flow.

## Deployment Instructions

### Option 1: Deploy via Azure Portal (Easiest)

1. Go to your Function App in Azure Portal
2. Click "Deployment Center" in the left menu
3. Choose deployment method:
   - **GitHub**: Connect your repository for automatic deployments
   - **Local Git**: Push directly from your machine
   - **ZIP Deploy**: Upload the `azure-function` folder as a zip

### Option 2: Deploy via VS Code

1. Install "Azure Functions" extension in VS Code
2. Sign in to Azure
3. Right-click the `azure-function` folder
4. Select "Deploy to Function App..."
5. Choose your Function App

### Option 3: Deploy via Azure CLI

```bash
cd azure-function
az login
az functionapp deployment source config-zip \
  --resource-group student-retention-rg \
  --name your-function-app-name \
  --src azure-function.zip
```

## Environment Variables

After deployment, configure these environment variables in Azure Portal:

1. Go to Function App → Configuration → Application settings
2. Add these variables:
   - `AZURE_TENANT_ID`: Your Azure AD tenant ID
   - `AZURE_CLIENT_ID`: Your app registration client ID
   - `AZURE_CLIENT_SECRET`: Your app client secret

## Testing

Endpoint: `https://your-function-app.azurewebsites.net/api/exchange-token`

Request:
```json
POST /api/exchange-token
Content-Type: application/json

{
  "token": "your-office-sso-token"
}
```

Response:
```json
{
  "accessToken": "graph-api-token",
  "expiresIn": 3600,
  "tokenType": "Bearer"
}
```
