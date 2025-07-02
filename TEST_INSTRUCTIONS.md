# Testing the Azure Function

This document explains how to test the `msGraphConnector` Azure Function locally using the provided test script.

## Prerequisites

1. **Azure AD App Registration**: You need an Azure AD app registration with the following configuration:
   - **Application Type**: Public client/native
   - **Redirect URI**: `http://localhost:3000/auth/callback` (Web)
   - **API Permissions**: 
     - Microsoft Graph: `Files.Read.All` (Application)
     - Microsoft Graph: `Sites.Read.All` (Application)
   - **Admin Consent**: Required for the above permissions

2. **Azure Function Running Locally**: The function must be running on your local machine
   ```bash
   npm start
   # or
   func start
   ```

3. **Environment Variables**: Set the following environment variables:
   ```bash
   export AZURE_CLIENT_ID=your-app-registration-client-id
   export AZURE_TENANT_ID=your-azure-tenant-id
   ```

## Running the Test

1. **Start the Azure Function locally**:
   ```bash
   npm start
   ```
   The function should be available at `http://localhost:7071/api/msGraphConnector`

2. **Run the test script**:
   ```bash
   npm run test-function
   # or
   npm run test-function "your search term"
   ```

3. **Authentication Flow**:
   - The script will automatically open your default browser
   - Sign in with your Microsoft account
   - Grant the requested permissions
   - The browser will redirect to a success page
   - Return to the terminal to see the test results

## What the Test Script Does

1. **Authentication**: Uses OAuth 2.0 authorization code flow to get a user token
2. **Token Exchange**: The Azure Function will exchange your user token for an application token using the "on-behalf-of" flow
3. **Function Test**: Makes a POST request to the local function with:
   - Authorization header containing the bearer token
   - Search term in the request body
4. **Results**: Displays the function response, including any found files

## Troubleshooting

### "ECONNREFUSED" Error
- Make sure the Azure Function is running locally (`npm start`)
- Verify the function URL is correct (default: `http://localhost:7071/api/msGraphConnector`)

### Authentication Errors
- Check that your Azure AD app registration has the correct permissions
- Ensure admin consent has been granted for the required Graph permissions
- Verify the redirect URI is configured correctly

### "No results found"
- The search term might not match any files in your OneDrive/SharePoint
- Try different search terms like "test", "document", or file extensions like ".docx"

### Permission Errors
- Ensure your Azure AD app has `Files.Read.All` and `Sites.Read.All` permissions
- Make sure these permissions have admin consent granted

## Customization

You can customize the test script by setting environment variables:

```bash
export AZURE_CLIENT_ID=your-client-id
export AZURE_TENANT_ID=your-tenant-id
export FUNCTION_URL=http://localhost:7071/api/msGraphConnector
npm run test-function
```

## Security Notes

- The test script only works with localhost redirect URIs for security
- Tokens are only stored in memory during the test execution
- No sensitive information is logged to the console