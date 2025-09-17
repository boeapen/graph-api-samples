# Microsoft Graph API Console Application

This workspace contains a C# console application that demonstrates Microsoft Graph API integration with MSAL authentication.

## Project Overview
- **Type**: C# Console Application (.NET)
- **Purpose**: Microsoft Graph API integration for email operations
- **Authentication**: Microsoft Authentication Library (MSAL) with client credentials flow
- **Target**: Microsoft Graph US Government cloud endpoint

## Key Features
- Client credentials authentication flow
- Email inbox reading from Exchange Online
- File attachment handling (upload/download)
- Support for large file attachments via upload sessions

## Dependencies
- Microsoft.Graph
- Microsoft.Identity.Client
- System.Net.Http

## Configuration Required
Before running the application, ensure you have:
1. Valid Azure AD application registration
2. Client ID, Client Secret, and Tenant ID configured
3. Appropriate Microsoft Graph permissions granted
4. Access to Microsoft Graph US Government endpoints

## Security Note
This application contains sensitive authentication credentials that should be secured appropriately in production environments.