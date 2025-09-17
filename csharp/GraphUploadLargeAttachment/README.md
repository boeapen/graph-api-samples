# Microsoft Graph API Console Application - Global Environment

A C# console application that demonstrates Microsoft Graph API integration with MSAL authentication for the global Microsoft 365 environment. This application creates draft emails with large attachments and sends them.

## Features

- **Authentication**: Microsoft Authentication Library (MSAL) with client credentials flow for Global Microsoft Graph
- **Draft Email Creation**: Creates draft messages programmatically
- **Large File Upload**: Uses createUploadSession API for uploading large attachments (>3MB)
- **Chunked Upload**: Implements optimized chunked upload strategy (3MB chunks for maximum efficiency)
- **Email Sending**: Sends the draft email with attachments
- **Global Environment**: Configured for global Microsoft Graph endpoints

## Prerequisites

- .NET 8.0 SDK or later
- Valid Azure AD application registration with appropriate permissions
- Access to Microsoft Graph US Government endpoints

## Required Azure AD Permissions

Your Azure AD application must have the following Microsoft Graph permissions:
- `Mail.ReadWrite` - To read and write mail
- `Mail.Send` - To send mail with attachments

## Configuration

Before running the application, update the following values in `Program.cs`:

```csharp
string ClientID = "your-client-id";
string ClientSecret = "your-client-secret";
string TenantID = "your-tenant-id";
string AuthorityURL = $"https://login.microsoftonline.com/{TenantID}";
string UserEmail = "sender@yourdomain.com"; // Replace with sender's email
string RecipientEmail = "recipient@yourdomain.com"; // Replace with recipient's email
string FilePath = @"C:\temp\large-file.pdf"; // Replace with actual file path
```

**Important Global Configuration Notes:**
- Use standard Microsoft 365 email addresses
- Ensure your Azure AD tenant is in the global Microsoft 365 environment
- Authority URL uses `login.microsoftonline.com`
- Graph base URL uses `graph.microsoft.com`

## Building and Running

### Build the project:
```bash
dotnet build
```

### Run the application:
```bash
dotnet run
```

## Project Structure

```
├── Program.cs              # Main application code
├── GraphApiApp.csproj      # Project file with NuGet dependencies
├── README.md              # This file
└── .github/
    └── copilot-instructions.md  # Workspace instructions
```

## Dependencies

- **Microsoft.Graph** (5.62.0) - Microsoft Graph SDK
- **Microsoft.Graph.Auth** (1.0.0-preview.7) - Authentication extensions
- **Microsoft.Identity.Client** (4.64.1) - Microsoft Authentication Library

## Application Flow

1. **Authentication**: Acquires access token using client credentials flow
2. **Email Reading**: Retrieves inbox messages with specified properties
3. **Upload Session**: Creates upload session for large file attachments
4. **Attachment Retrieval**: Downloads attachments from specified messages

## Security Considerations

⚠️ **Important Security Notes:**
- Never commit credentials to source control
- Store sensitive configuration in Azure Key Vault or environment variables
- Use managed identities when running in Azure
- Rotate client secrets regularly

## Error Handling

The application includes comprehensive error handling for:
- MSAL authentication exceptions
- HTTP request failures
- File system operations
- Network connectivity issues

## Troubleshooting

### Common Issues:

1. **Authentication Failures**
   - Verify client ID, secret, and tenant ID are correct
   - Ensure the Azure AD app has required permissions
   - Check if permissions are granted admin consent

2. **Graph API Errors**
   - Verify the user exists and has a mailbox
   - Check message IDs are valid and accessible
   - Ensure proper scopes are configured

3. **File Operations**
   - Verify file paths exist for upload operations
   - Check file permissions and sizes
   - Ensure network connectivity for downloads

## Development

To modify or extend this application:
1. Update the project file to add new dependencies
2. Modify the authentication scope if accessing different Graph APIs
3. Add error logging for production environments
4. Implement configuration file or environment variable support

## License

This project is for demonstration purposes. Please ensure compliance with your organization's security policies before use.