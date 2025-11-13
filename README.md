# Squish the Phish - Outlook Add-in

A prototype Outlook add-in that provides a custom "Report Phish or Spam" button, allowing users to submit suspicious emails directly to Microsoft security services.

## Overview

This add-in uses:
- **Entra ID SSO with Nested App Authentication (NAA)** for secure authentication
- **Microsoft Graph API emailThreatSubmission** to report phishing and spam emails
- Fallback to Office dialog API when NAA is unavailable

## Prerequisites

- Microsoft 365 subscription
- [Node.js](https://nodejs.org/) (latest recommended)
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8+
- Azure subscription (for deployment)

## Setup

### 1. Register Entra ID Application

1. Go to [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908)
2. Sign in with **admin credentials**
3. Select **New registration**:
   - **Name**: `Squish-the-Phish`
   - **Supported account types**: Accounts in any organizational directory (Multitenant) and personal Microsoft accounts
   - **Redirect URI**: Select **Single-page application (SPA)** and enter `brk-multihub://localhost:3000`
4. Copy the **Application (client) ID**
5. Under **Authentication**, add these redirect URIs:
   - `https://localhost:3000/auth.html`
   - `https://localhost:3000/dialog.html`
6. Under **API permissions**, grant:
   - `ThreatSubmission.ReadWrite` (for Graph API emailThreatSubmission)
   - `User.Read` (for user profile)
7. After the add-in deployed to Azure Static Web App, add ASWA url to steps 3 and 5.

### 2. Configure the Sample

1. Clone this repository
2. Open `src/spamreporting/msalconfig.ts`
3. Replace `"Enter_the_Application_Id_Here"` with your Application ID
4. Save the file

### 3. Run Locally

```bash
npm install
npm run start
```

The add-in will sideload into Outlook. Open a message and select the **Report Phish** button to test.

## Deployment to Azure Static Web Apps

### Setup Azure Static Web App

1. Create a Static Web App in [Azure Portal](https://portal.azure.com)
2. Copy the deployment token from **Settings** > **Configuration**
3. Add the token as a GitHub secret named `AZURE_STATIC_WEB_APPS_API_TOKEN`
4. Update the production URL in `webpack.config.js`:
   ```javascript
   const urlProd = "https://YOUR_APP_NAME.azurestaticapps.net/";
   ```
5. Push to `main` branch - GitHub Actions will automatically build and deploy

### Configuration Files

- `staticwebapp.config.json` - Azure routing and CORS configuration
- `.github/workflows/azure-static-web-apps-deploy.yml` - CI/CD pipeline

## User Permissions

Users must have the following permissions to submit emails:
- Member of organization with Entra ID access
- Granted consent to the application's required API permissions (Mail.Read, User.Read)
- Access to Microsoft Graph API emailThreatSubmission endpoint (ThreatSubmission.ReadWrite)

## Key Implementation

- `src/spamreporting/authConfig.ts` - MSAL NAA configuration and token management
- `src/spamreporting/msgraph-helper.ts` - Microsoft Graph API calls
- `src/spamreporting/spamreporting.ts` - UI interaction handlers

## Resources

- [NAA Documentation](https://aka.ms/NAAdocs)
- [NAA FAQ](https://aka.ms/NAAFAQ)
- [Graph API emailThreatSubmission](https://learn.microsoft.com/graph/api/resources/security-emailthreatsubmission)


## License
Copyright (c) 2025 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
