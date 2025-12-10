using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace PBNetLibrary
{
    /// <summary>
    /// Simple Email Service for PowerBuilder
    /// </summary>
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-4A5B-8C9D-0E1F2A3B4C5D")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class SendEmail
    {
        private string _logFilePath;
        private StringBuilder _detailedLog;

        public SendEmail()
        {
            _detailedLog = new StringBuilder();
            // Log files are written to Windows TEMP directory
            string tempPath = Path.GetTempPath();
            _logFilePath = Path.Combine(tempPath, $"PB_Email_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            WriteLog($"Log file created: {_logFilePath}");
        }

        private void WriteLog(string message)
        {
            string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            _detailedLog.AppendLine(logEntry);
            
            try
            {
                File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Ignore file write errors
            }
        }

        /// <summary>
        /// Get the path to the current log file
        /// </summary>
        public string GetLogFilePath()
        {
            return _logFilePath;
        }

        /// <summary>
        /// Get the detailed log as a string
        /// </summary>
        public string GetDetailedLog()
        {
            return _detailedLog.ToString();
        }

        /// <summary>
        /// Clear the current log
        /// </summary>
        public void ClearLog()
        {
            _detailedLog.Clear();
        }

        /// <summary>
        /// Universal Email Send Method - Supports all cases (Returns: 1=Success, 0=Failed)
        /// - Body: Plain text or HTML (auto-detected)
        /// - TO: Single email or multiple (semicolon-separated: "email1@test.com;email2@test.com")
        /// - CC: Single email or multiple (semicolon-separated) - use empty string "" for none
        /// - BCC: Single email or multiple (semicolon-separated) - use empty string "" for none
        /// - Attachments: Single file or multiple (semicolon-separated paths) - use empty string "" for none
        /// </summary>
        public int Send(
            string smtpServer,
            int port,
            string username,
            string password,
            string senderName,
            string senderEmail,
            string toEmails,
            string ccEmails,
            string bccEmails,
            string subject,
            string body,
            string attachmentPaths)
        {
            try
            {
                WriteLog("=== Starting Email Send ===");
                WriteLog($"SMTP: {smtpServer}:{port}");
                WriteLog($"From: {senderName} <{senderEmail}>");
                                                WriteLog($"Subject: {subject}");

                // Validate inputs
                if (string.IsNullOrWhiteSpace(smtpServer))
                {
                    WriteLog("ERROR: SMTP server is empty");
                    return 0;
                }
                if (string.IsNullOrWhiteSpace(username))
                {
                    WriteLog("ERROR: Username is empty");
                    return 0;
                }
                if (string.IsNullOrWhiteSpace(toEmails))
                {
                    WriteLog("ERROR: Recipient email(s) are empty");
                    return 0;
                }

                // Create message
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress(senderName, senderEmail));
                message.Subject = subject;

                // Parse TO recipients
                string[] toList = toEmails.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                int toCount = 0;
                foreach (string email in toList)
                {
                    string trimmedEmail = email.Trim();
                    if (!string.IsNullOrWhiteSpace(trimmedEmail))
                    {
                        message.To.Add(new MailboxAddress(trimmedEmail, trimmedEmail));
                        WriteLog($"TO: {trimmedEmail}");
                        toCount++;
                    }
                }

                if (toCount == 0)
                {
                    WriteLog("ERROR: No valid TO recipients");
                    return 0;
                }

                // Parse CC recipients (optional)
                if (!string.IsNullOrWhiteSpace(ccEmails))
                {
                    string[] ccList = ccEmails.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string email in ccList)
                    {
                        string trimmedEmail = email.Trim();
                        if (!string.IsNullOrWhiteSpace(trimmedEmail))
                        {
                            message.Cc.Add(new MailboxAddress(trimmedEmail, trimmedEmail));
                            WriteLog($"CC: {trimmedEmail}");
                        }
                    }
                }

                // Parse BCC recipients (optional)
                if (!string.IsNullOrWhiteSpace(bccEmails))
                {
                    string[] bccList = bccEmails.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string email in bccList)
                    {
                        string trimmedEmail = email.Trim();
                        if (!string.IsNullOrWhiteSpace(trimmedEmail))
                        {
                            message.Bcc.Add(new MailboxAddress(trimmedEmail, trimmedEmail));
                            WriteLog($"BCC: {trimmedEmail}");
                        }
                    }
                }

                // Build body - auto-detect HTML or plain text
                var bodyBuilder = new BodyBuilder();
                if (!string.IsNullOrWhiteSpace(body))
                {
                    // Check if body contains HTML
                    if (body.TrimStart().StartsWith("<") || body.Contains("<html") || body.Contains("<HTML") || 
                        body.Contains("<body") || body.Contains("<BODY") || body.Contains("<p>") || body.Contains("<div>") ||
                        body.Contains("<br>") || body.Contains("<BR>") || body.Contains("<tr>") || body.Contains("<TR>") || 
                        body.Contains("<strong>") || body.Contains("<STRONG>"))
                    {
                        bodyBuilder.HtmlBody = body;
                        // Also provide plain text version by stripping HTML tags
                        bodyBuilder.TextBody = System.Text.RegularExpressions.Regex.Replace(body, "<.*?>", "");
                        WriteLog("Body format: HTML with plain text fallback");
                    }
                    else
                    {
                        bodyBuilder.TextBody = body;
                        WriteLog("Body format: Plain Text");
                    }
                }

                // Parse and add attachments (optional)
                if (!string.IsNullOrWhiteSpace(attachmentPaths))
                {
                    string[] pathList = attachmentPaths.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                    int attachmentCount = 0;
                    foreach (string path in pathList)
                    {
                        string trimmedPath = path.Trim();
                        if (!string.IsNullOrWhiteSpace(trimmedPath))
                        {
                            if (File.Exists(trimmedPath))
                            {
                                bodyBuilder.Attachments.Add(trimmedPath);
                                var fileInfo = new FileInfo(trimmedPath);
                                WriteLog($"Attached: {Path.GetFileName(trimmedPath)} ({fileInfo.Length / 1024.0:F2} KB)");
                                attachmentCount++;
                            }
                            else
                            {
                                WriteLog($"WARNING: Attachment not found: {trimmedPath}");
                            }
                        }
                    }
                    if (attachmentCount > 0)
                    {
                        WriteLog($"Total attachments: {attachmentCount}");
                    }
                }

                message.Body = bodyBuilder.ToMessageBody();

                // Send email
                using (var client = new SmtpClient())
                {
                    client.Connect(smtpServer, port, SecureSocketOptions.StartTls);
                    client.Authenticate(username, password);
                    client.Send(message);
                    client.Disconnect(true);
                }

                WriteLog($"SUCCESS: Email sent to {toCount} recipient(s)");
                WriteLog("=== End Email Send ===");
                return 1;
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR: {ex.Message}");
                WriteLog("=== End Email Send ===");
                return 0;
            }
        }

        /// <summary>
        /// Get last error message
        /// </summary>
        public string GetLastError()
        {
            string log = _detailedLog.ToString();
            string[] lines = log.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            
            for (int i = lines.Length - 1; i >= 0; i--)
            {
                if (lines[i].Contains("ERROR:"))
                {
                    return lines[i];
                }
            }
            
            return "No error found";
        }
    }
}
