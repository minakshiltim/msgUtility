using System;
using System.IO;
using System.Text;
// Namespaces live in Mail.dll
// using Limilabs.Mail;
// using Limilabs.Mail.Fluent;
// Licensing namespace can be inside Mail.dll in newer builds
// If present, uncomment and use it:
// using Limilabs.Licensing;

//Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

using MimeKit;
using MsgKit;
using MsgKit.Enums;

internal static class Program
{
    // Usage:
    //   dotnet run -- "C:\in\sample.eml" "C:\out\sample.msg"
    private static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string inputEml = "C:\\Users\\10844924\\Documents\\Workspace\\Project\\Sample\\emlFormatEmail\\FW__Reinsurance_Report___Insured__New_Mexico_Counties_Insurance_Authority_Multi_Line_Program_____Claimant___Estate_of_Debbie_Madrid___BPE_Claim___2105547______Policy_PEM0000144_01__.eml";
        string outputMsg = "C:\\Users\\10844924\\Documents\\Workspace\\Project\\Sample\\emlFormatEmail\\Output\\limilabs";

        var tempFiles = new List<string>();

        try
        {
            var mime = MimeMessage.Load(inputEml);
            var from = mime.From.Mailboxes.FirstOrDefault();
            var subject = mime.Subject ?? "(no subject)";
            var msgSender = new MsgKit.Sender(
                from?.Address ?? "unknown@example.com",
                from?.Name ?? "Unknown"
            );
            
            using var email = new Email(msgSender, null, subject);

            // Set body content
            email.BodyText = mime.TextBody ?? "";
            email.BodyHtml = mime.HtmlBody ?? "";

            // Add recipients
            foreach (var to in mime.To.Mailboxes)
                email.Recipients.AddTo(to.Address, to.Name);
            foreach (var cc in mime.Cc.Mailboxes)
                email.Recipients.AddCc(cc.Address, cc.Name);
            foreach (var bcc in mime.Bcc.Mailboxes)
                email.Recipients.AddBcc(bcc.Address, bcc.Name);

            // Set sent and received times for MAPI properties
            if (mime.Date != DateTimeOffset.MinValue)
            {
                email.SentOn = mime.Date.DateTime;
                email.ReceivedOn = mime.Date.DateTime;
            }

            // Set importance/priority
            if (mime.Importance == MimeKit.MessageImportance.High)
                email.Importance = MsgKit.Enums.MessageImportance.IMPORTANCE_HIGH;
            else if (mime.Importance == MimeKit.MessageImportance.Low)
                email.Importance = MsgKit.Enums.MessageImportance.IMPORTANCE_LOW;
            else
                email.Importance = MsgKit.Enums.MessageImportance.IMPORTANCE_NORMAL;

            // Process all attachments including nested emails
            ProcessAttachments(mime, email, tempFiles);

            // Ensure output directory exists
            Directory.CreateDirectory(outputMsg);

            var baseName = Path.GetFileNameWithoutExtension(inputEml);
            var msgPath = Path.Combine(outputMsg, baseName + ".msg");
            email.Save(msgPath);

            Console.WriteLine($"Converted OK → {msgPath}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Conversion failed:");
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
        finally
        {
            // Clean up temp files and directories after email is saved
            foreach (var tempFile in tempFiles)
            {
                try
                {
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                    else if (Directory.Exists(tempFile))
                        Directory.Delete(tempFile, true);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }

    private static void ProcessAttachments(MimeMessage mime, Email email, List<string> tempFiles)
    {
        foreach (var attachment in mime.Attachments)
        {
            if (attachment is MessagePart messagePart)
            {
                // Handle nested email attachments (message/rfc822) - convert to MSG format
                var nestedMessage = messagePart.Message;
                var attachmentName = messagePart.ContentDisposition?.FileName ?? 
                                    nestedMessage.Subject ?? 
                                    $"email_{Guid.NewGuid()}";
                
                // Remove .eml extension if present and ensure .msg extension
                if (attachmentName.EndsWith(".eml", StringComparison.OrdinalIgnoreCase))
                    attachmentName = Path.GetFileNameWithoutExtension(attachmentName);
                if (!attachmentName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                    attachmentName += ".msg";

                // Create unique temp directory to avoid file conflicts while preserving original name
                var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);
                var tempMsgPath = Path.Combine(tempDir, attachmentName);
                
                // Convert nested email to MSG format
                var nestedMsgEmail = ConvertMimeMessageToMsg(nestedMessage, tempFiles);
                nestedMsgEmail.Save(tempMsgPath);
                nestedMsgEmail.Dispose();
                
                email.Attachments.Add(tempMsgPath);
                tempFiles.Add(tempMsgPath);
                tempFiles.Add(tempDir); // Add directory for cleanup
            }
            else if (attachment is MimePart mimePart)
            {
                // Handle regular file attachments
                var attachmentName = mimePart.FileName ?? 
                                    $"attachment_{Guid.NewGuid()}{GetExtensionFromMimeType(mimePart.ContentType.MimeType)}";
                
                // Create unique temp directory to avoid file conflicts while preserving original name
                var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);
                var tempPath = Path.Combine(tempDir, attachmentName);
                
                // Use Stream property to get raw decoded content without any text transformations
                using (var contentStream = mimePart.Content.Open())
                using (var fileStream = File.Create(tempPath))
                {
                    contentStream.CopyTo(fileStream);
                }
                
                email.Attachments.Add(tempPath);
                tempFiles.Add(tempPath);
                tempFiles.Add(tempDir); // Add directory for cleanup
            }
        }

        // Process inline attachments (images in signatures, HTML content, etc.)
        foreach (var bodyPart in mime.BodyParts.OfType<MimePart>())
        {
            // Check if it's an inline part (has ContentId) or is marked as attachment but not in main attachments list
            bool isInlineImage = !string.IsNullOrEmpty(bodyPart.ContentId) && 
                                 bodyPart.ContentDisposition?.Disposition == "inline";
            bool isUnprocessedAttachment = bodyPart.IsAttachment && !mime.Attachments.Contains(bodyPart);
            
            if (isInlineImage || isUnprocessedAttachment)
            {
                var attachmentName = bodyPart.FileName ?? 
                                    bodyPart.ContentId?.Trim('<', '>') ?? 
                                    $"inline_{Guid.NewGuid()}{GetExtensionFromMimeType(bodyPart.ContentType.MimeType)}";
                
                // Ensure proper extension
                if (string.IsNullOrEmpty(Path.GetExtension(attachmentName)))
                {
                    attachmentName += GetExtensionFromMimeType(bodyPart.ContentType.MimeType);
                }
                
                // Create unique temp directory to avoid file conflicts while preserving original name
                var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);
                var tempPath = Path.Combine(tempDir, attachmentName);
                
                // Use Stream property to get raw decoded content without any text transformations
                using (var contentStream = bodyPart.Content.Open())
                using (var fileStream = File.Create(tempPath))
                {
                    contentStream.CopyTo(fileStream);
                }
                
                // Add as inline attachment if it has a ContentId
                if (isInlineImage)
                {
                    email.Attachments.Add(tempPath, -1, true, bodyPart.ContentId?.Trim('<', '>'));
                }
                else
                {
                    email.Attachments.Add(tempPath);
                }
                
                tempFiles.Add(tempPath);
                tempFiles.Add(tempDir); // Add directory for cleanup
            }
        }
    }

    private static Email ConvertMimeMessageToMsg(MimeMessage mime, List<string> tempFiles)
    {
        var from = mime.From.Mailboxes.FirstOrDefault();
        var subject = mime.Subject ?? "(no subject)";
        var msgSender = new MsgKit.Sender(
            from?.Address ?? "unknown@example.com",
            from?.Name ?? "Unknown"
        );
        
        var email = new Email(msgSender, null, subject);

        // Set body content
        email.BodyText = mime.TextBody ?? "";
        email.BodyHtml = mime.HtmlBody ?? "";

        // Add recipients
        foreach (var to in mime.To.Mailboxes)
            email.Recipients.AddTo(to.Address, to.Name);
        foreach (var cc in mime.Cc.Mailboxes)
            email.Recipients.AddCc(cc.Address, cc.Name);
        foreach (var bcc in mime.Bcc.Mailboxes)
            email.Recipients.AddBcc(bcc.Address, bcc.Name);

        // Set sent and received times for MAPI properties
        if (mime.Date != DateTimeOffset.MinValue)
        {
            email.SentOn = mime.Date.DateTime;
            email.ReceivedOn = mime.Date.DateTime;
        }

        // Set importance/priority
        if (mime.Importance == MimeKit.MessageImportance.High)
            email.Importance = MsgKit.Enums.MessageImportance.IMPORTANCE_HIGH;
        else if (mime.Importance == MimeKit.MessageImportance.Low)
            email.Importance = MsgKit.Enums.MessageImportance.IMPORTANCE_LOW;
        else
            email.Importance = MsgKit.Enums.MessageImportance.IMPORTANCE_NORMAL;

        // Process all attachments including nested emails (recursively)
        ProcessAttachments(mime, email, tempFiles);

        return email;
    }

    private static string GetExtensionFromMimeType(string mimeType)
    {
        return mimeType?.ToLower() switch
        {
            "application/pdf" => ".pdf",
            "application/msword" => ".doc",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => ".docx",
            "application/vnd.ms-excel" => ".xls",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => ".xlsx",
            "application/zip" => ".zip",
            "text/plain" => ".txt",
            "text/html" => ".html",
            "image/jpeg" => ".jpg",
            "image/png" => ".png",
            "image/gif" => ".gif",
            _ => ".dat"
        };
    }
}