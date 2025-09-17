using Microsoft.Identity.Client;
using System.Text;
using System.Text.Json;

Console.WriteLine("Microsoft Graph API - Global Environment");
Console.WriteLine("Draft Email with Large Attachment Upload");
Console.WriteLine("======================================");

// Global Microsoft Graph Configuration
string ClientID = "your-client-id-here";
string ClientSecret = "your-client-secret-here";
string TenantID = "your-tenant-id-here";
string AuthorityURL = $"https://login.microsoftonline.com/{TenantID}";
string GraphBaseUrl = "https://graph.microsoft.com";
string AuthScopes = "https://graph.microsoft.com/.default";
string UserEmail = "sender@yourdomain.com"; // Replace with actual sender email
string RecipientEmail = "recipient@yourdomain.com"; // Replace with actual recipient email

// File to upload (replace with actual file path)
string FilePath = @"C:\path\to\your\large-file.pdf";

// Initialize MSAL client for Global Microsoft Graph
IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(ClientID)
    .WithClientSecret(ClientSecret)
    .WithTenantId(TenantID)
    .WithAuthority(AuthorityURL)
    .Build();

string[] scopes = AuthScopes.Split(',');
AuthenticationResult? tokenResult = null;
string accessToken = "";

// Step 1: Acquire access token
try
{
    Console.WriteLine("Step 1: Acquiring access token for Global Microsoft Graph...");
    tokenResult = await app.AcquireTokenForClient(scopes).ExecuteAsync();
    accessToken = tokenResult.AccessToken;
    Console.WriteLine($"✓ Token acquired successfully. Expires: {tokenResult.ExpiresOn.UtcDateTime}");
}
catch (MsalServiceException ex)
{
    Console.WriteLine($"❌ MSAL Service Exception: {ex.Message}");
    Console.WriteLine("❌ Authentication failed. Please check your credentials and try again.");
    return;
}
catch (Exception ex)
{
    Console.WriteLine($"❌ Authentication Exception: {ex.Message}");
    return;
}

using (var client = new HttpClient())
{
    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
    client.DefaultRequestHeaders.Add("Accept", "application/json");

    string draftMessageId = "";

    // Step 2: Create draft message
    try
    {
        Console.WriteLine("\nStep 2: Creating draft message...");
        
        var draftMessage = new
        {
            subject = "Test Email with Large Attachment - Global Graph API",
            body = new
            {
                contentType = "HTML",
                content = @"
                    <html>
                    <body>
                        <h2>Test Email from Global Microsoft Graph API</h2>
                        <p>This email was created using Microsoft Graph API in the global Microsoft 365 environment.</p>
                        <p>This message includes a large file attachment uploaded via createUploadSession API.</p>
                        <p><strong>Timestamp:</strong> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + @"</p>
                    </body>
                    </html>"
            },
            toRecipients = new[]
            {
                new
                {
                    emailAddress = new
                    {
                        address = RecipientEmail,
                        name = "Test Recipient"
                    }
                }
            }
        };

        var jsonContent = JsonSerializer.Serialize(draftMessage, new JsonSerializerOptions { WriteIndented = true });
        var draftContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

        string createDraftUrl = $"{GraphBaseUrl}/beta/users/{UserEmail}/messages";
        HttpResponseMessage draftResponse = await client.PostAsync(createDraftUrl, draftContent);

        if (draftResponse.IsSuccessStatusCode)
        {
            string draftResponseContent = await draftResponse.Content.ReadAsStringAsync();
            var draftResult = JsonSerializer.Deserialize<JsonElement>(draftResponseContent);
            draftMessageId = draftResult.GetProperty("id").GetString()!;
            Console.WriteLine($"✓ Draft message created successfully. ID: {draftMessageId}");
        }
        else
        {
            string errorContent = await draftResponse.Content.ReadAsStringAsync();
            Console.WriteLine($"❌ Failed to create draft message: {draftResponse.StatusCode}");
            Console.WriteLine($"Error details: {errorContent}");
            return;
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Draft creation exception: {ex.Message}");
        return;
    }

    // Step 3: Create upload session for large attachment
    try
    {
        Console.WriteLine("\nStep 3: Creating upload session for large attachment...");

        if (!File.Exists(FilePath))
        {
            Console.WriteLine($"❌ File not found: {FilePath}");
            Console.WriteLine("Please update the FilePath variable with a valid file path.");
            return;
        }

        FileInfo fileInfo = new FileInfo(FilePath);
        Console.WriteLine($"File to upload: {fileInfo.Name} ({fileInfo.Length:N0} bytes)");

        var uploadSessionRequest = new
        {
            AttachmentItem = new
            {
                attachmentType = "file",
                name = fileInfo.Name,
                size = fileInfo.Length,
                contentType = "application/octet-stream"
            }
        };

        var uploadSessionJson = JsonSerializer.Serialize(uploadSessionRequest, new JsonSerializerOptions { WriteIndented = true });
        var uploadSessionContent = new StringContent(uploadSessionJson, Encoding.UTF8, "application/json");

        string createUploadSessionUrl = $"{GraphBaseUrl}/beta/users/{UserEmail}/messages/{draftMessageId}/attachments/createUploadSession";
        HttpResponseMessage uploadSessionResponse = await client.PostAsync(createUploadSessionUrl, uploadSessionContent);

        if (uploadSessionResponse.IsSuccessStatusCode)
        {
            string uploadSessionResponseContent = await uploadSessionResponse.Content.ReadAsStringAsync();
            var uploadSessionResult = JsonSerializer.Deserialize<JsonElement>(uploadSessionResponseContent);
            string uploadUrl = uploadSessionResult.GetProperty("uploadUrl").GetString()!;
            
            Console.WriteLine("✓ Upload session created successfully.");
            Console.WriteLine($"Upload URL: {uploadUrl}");

            // Step 4: Upload file in chunks
            Console.WriteLine("\nStep 4: Uploading file in chunks...");
            
            const int chunkSize = 3 * 1024 * 1024; // 3MB chunks (maximum supported by Graph API)
            byte[] fileBytes = await File.ReadAllBytesAsync(FilePath);
            long totalSize = fileBytes.Length;
            long uploadedBytes = 0;

            using (var uploadClient = new HttpClient())
            {
                while (uploadedBytes < totalSize)
                {
                    long remainingBytes = totalSize - uploadedBytes;
                    int currentChunkSize = (int)Math.Min(chunkSize, remainingBytes);
                    
                    byte[] chunk = new byte[currentChunkSize];
                    Array.Copy(fileBytes, uploadedBytes, chunk, 0, currentChunkSize);

                    var chunkContent = new ByteArrayContent(chunk);
                    chunkContent.Headers.Add("Content-Range", 
                        $"bytes {uploadedBytes}-{uploadedBytes + currentChunkSize - 1}/{totalSize}");

                    HttpResponseMessage chunkResponse = await uploadClient.PutAsync(uploadUrl, chunkContent);
                    
                    if (chunkResponse.IsSuccessStatusCode || chunkResponse.StatusCode == System.Net.HttpStatusCode.Created)
                    {
                        uploadedBytes += currentChunkSize;
                        double progressPercentage = (double)uploadedBytes / totalSize * 100;
                        Console.WriteLine($"  Uploaded: {uploadedBytes:N0}/{totalSize:N0} bytes ({progressPercentage:F1}%)");
                    }
                    else
                    {
                        string chunkError = await chunkResponse.Content.ReadAsStringAsync();
                        Console.WriteLine($"❌ Chunk upload failed: {chunkResponse.StatusCode}");
                        Console.WriteLine($"Error details: {chunkError}");
                        return;
                    }
                }
            }

            Console.WriteLine("✓ File upload completed successfully!");
        }
        else
        {
            string uploadSessionError = await uploadSessionResponse.Content.ReadAsStringAsync();
            Console.WriteLine($"❌ Failed to create upload session: {uploadSessionResponse.StatusCode}");
            Console.WriteLine($"Error details: {uploadSessionError}");
            return;
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Upload session exception: {ex.Message}");
        return;
    }

    // Step 5: Send the draft email
    try
    {
        Console.WriteLine("\nStep 5: Sending the draft email...");

        string sendDraftUrl = $"{GraphBaseUrl}/beta/users/{UserEmail}/messages/{draftMessageId}/send";
        var sendContent = new StringContent("{}", Encoding.UTF8, "application/json");

        HttpResponseMessage sendResponse = await client.PostAsync(sendDraftUrl, sendContent);

        if (sendResponse.IsSuccessStatusCode)
        {
            Console.WriteLine("✓ Email sent successfully!");
            Console.WriteLine($"📧 Draft message with ID {draftMessageId} has been sent with the large attachment.");
        }
        else
        {
            string sendError = await sendResponse.Content.ReadAsStringAsync();
            Console.WriteLine($"❌ Failed to send email: {sendResponse.StatusCode}");
            Console.WriteLine($"Error details: {sendError}");
            return;
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Send email exception: {ex.Message}");
        return;
    }
}

Console.WriteLine("\n🎉 All operations completed successfully!");
Console.WriteLine("\nWorkflow Summary:");
Console.WriteLine("1. ✓ Authenticated with Global Microsoft Graph API");
Console.WriteLine("2. ✓ Created draft message");
Console.WriteLine("3. ✓ Created upload session for large attachment");
Console.WriteLine("4. ✓ Uploaded file in chunks");
Console.WriteLine("5. ✓ Sent the email with attachment");
Console.WriteLine("\nPress any key to exit...");
Console.ReadKey();
