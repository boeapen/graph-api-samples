# Configuration Template

Copy this template and update with your actual values before running the application.

## Azure AD Application Configuration

```csharp
// Replace these values in Program.cs:
string ClientID = "your-azure-ad-app-client-id";
string ClientSecret = "your-azure-ad-app-client-secret";
string TenantID = "your-azure-ad-tenant-id";
```

## Email Configuration

```csharp
// Replace these values in Program.cs:
string UserEmail = "sender@yourdomain.com";        // Mailbox owner (sender)
string RecipientEmail = "recipient@yourdomain.com"; // Email recipient
```

## File Upload Configuration

```csharp
// Replace this value in Program.cs:
string FilePath = @"C:\path\to\your\large-file.pdf"; // File to attach
```

## Required Azure AD Permissions

Your Azure AD application must have the following Microsoft Graph permissions:
- `Mail.ReadWrite` - To create and modify messages
- `Mail.Send` - To send messages

## Setup Instructions

1. **Create Azure AD App Registration:**
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to Azure Active Directory > App registrations
   - Click "New registration"
   - Note the Application (client) ID and Directory (tenant) ID

2. **Create Client Secret:**
   - In your app registration, go to "Certificates & secrets"
   - Click "New client secret"
   - Copy the secret value immediately (it won't be shown again)

3. **Grant Permissions:**
   - Go to "API permissions"
   - Add Microsoft Graph permissions: `Mail.ReadWrite` and `Mail.Send`
   - Grant admin consent

4. **Update Program.cs:**
   - Replace all placeholder values with your actual configuration
   - Ensure file path points to an existing file

## Security Notes

⚠️ **Never commit actual credentials to source control!**
- Use environment variables in production
- Consider Azure Key Vault for storing secrets
- Rotate client secrets regularly