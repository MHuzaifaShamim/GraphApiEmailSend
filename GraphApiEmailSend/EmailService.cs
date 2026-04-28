
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Messages.Item.Attachments.CreateUploadSession;

public class EmailService
{
    private readonly GraphServiceClient _graphClient;

    public EmailService(string tenantId, string clientId, string clientSecret)
    {
        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        _graphClient = new GraphServiceClient(credential);
    }
    // R&D Task: Testing PR workflow).
    // R&D Task: Testing PR workflow).
    public async Task SendProfessionalEmailAsync(string senderEmail, string recipientEmail, string recipientName, List<string> filePaths)
    {
        try
        {
            var draftMessage = new Message
            {
                Subject = "📢 Official Update: AwesomeTech Inc.",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = GetHtmlTemplate(recipientName)
                },
                ToRecipients = new List<Recipient>
                {
                    new Recipient { EmailAddress = new EmailAddress { Address = recipientEmail } }
                }
            };

            var savedDraft = await _graphClient.Users[senderEmail].Messages.PostAsync(draftMessage);
            string messageId = savedDraft.Id;

            Console.WriteLine($"[Step 1] Draft created. Processing attachments...");

            bool hasAttachments = false;
            foreach (var filePath in filePaths)
            {
                if (File.Exists(filePath))
                {
                    await UploadLargeFileAsync(senderEmail, messageId, filePath);
                    hasAttachments = true;
                }
                else
                {
                    Console.WriteLine($"⚠️ Warning: File not found: {filePath}");
                }
            }

            if (!hasAttachments)
            {
                Console.WriteLine("ℹ️ Info: No valid files found to attach. Sending email without attachments.");
            }

            await _graphClient.Users[senderEmail].Messages[messageId].Send.PostAsync();

            Console.WriteLine($"✅ Success: Email sent to {recipientName} with {filePaths.Count} file(s) processed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }

    private async Task UploadLargeFileAsync(string userId, string messageId, string filePath)
    {
        FileInfo fileInfo = new FileInfo(filePath);
        string contentType = GetContentType(fileInfo.Extension); 

        using var stream = File.OpenRead(filePath);

        var uploadSessionBody = new CreateUploadSessionPostRequestBody
        {
            AttachmentItem = new AttachmentItem
            {
                AttachmentType = AttachmentType.File,
                Name = fileInfo.Name,
                Size = fileInfo.Length,
                ContentType = contentType
            }
        };

        var uploadSession = await _graphClient.Users[userId]
            .Messages[messageId]
            .Attachments
            .CreateUploadSession
            .PostAsync(uploadSessionBody);

        int maxSliceSize = 320 * 1024;
        var fileUploadTask = new LargeFileUploadTask<FileAttachment>(uploadSession, stream, maxSliceSize, _graphClient.RequestAdapter);

        await fileUploadTask.UploadAsync();
        Console.WriteLine($"📎 Attached ({fileInfo.Extension.ToUpper()}): {fileInfo.Name}");
    }

    private string GetContentType(string extension)
    {
        return extension.ToLower() switch
        {
            ".pdf" => "application/pdf",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xls" => "application/vnd.ms-excel",
            _ => "application/octet-stream",
        };
    }

    private string GetHtmlTemplate(string recipientName)
    {
        return $@"
        <html>
        <body style='font-family: Arial, sans-serif; background-color: #f4f4f4; padding: 20px;'>
            <div style='max-width: 600px; margin: auto; border: 1px solid #ddd; padding: 20px; border-radius: 10px; background-color: #fff;'>
                <h2 style='color: #0078d4;'>Assalam-o-Alaikum, {recipientName}!</h2>
                <p>Please find the requested official documents attached to this email.</p>
                <div style='padding: 10px; border-left: 4px solid #0078d4; background: #f9f9f9;'>
                    <strong>Note:</strong> We have shared the latest Excel reports and PDF guides.
                </div>
                <p>Regards,<br><strong>AwesomeTech Inc.</strong></p>
            </div>
        </body>
        </html>";
    }
}