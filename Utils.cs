using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;

namespace MailCollector
{
    class Utils
    {
        const string UUID_PATTERN = "[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}";

        // https://stackoverflow.com/questions/29912136/converting-a-outlook-mail-attachment-to-byte-array-with-c-sharp
        const string PR_ATTACH_DATA_BIN = "http://schemas.microsoft.com/mapi/proptag/0x37010102";

        static readonly Regex UuidRegex = new Regex(UUID_PATTERN);

        public enum DeliveryType
        {
            AS_LINK,
            AS_ATTACHMENT,
            AS_BODY
        }

        public static string NormalisePath(string tempPath)
        {
            string path = Path.GetFullPath(tempPath);
            if (!Directory.Exists(path))
            {
                Console.WriteLine("[+] Payload save directory doesn't exist, attempting to create it.");
                Directory.CreateDirectory(path);
            }
            return path;
        }

        public static string CreateUniqueTempDirectory(string path = "")
        {
            string basePath = string.IsNullOrEmpty(path) ? Path.GetTempPath() : path;
            string uniqueTempDir = Path.GetFullPath(Path.Combine(basePath, Guid.NewGuid().ToString()));

            Directory.CreateDirectory(uniqueTempDir);
            return uniqueTempDir;
        }

        public static void SaveMsg(MailItem mailItem, string tempPath, string emailId)
        {
            string msgFileName = Path.Combine(tempPath, $"{emailId}.msg");
            mailItem.SaveAs(msgFileName, OlSaveAsType.olMSG);
        }

        // Resolves an address entry to an SMTP address. Exchange entries expose this via
        // GetExchangeUser(); non-Exchange entries (e.g. POP/IMAP/Outlook.com) return null
        // there but carry the SMTP address directly in Address.
        public static string GetSmtpAddress(AddressEntry entry)
        {
            if (entry == null)
                return "";

            try
            {
                ExchangeUser exchangeUser = entry.GetExchangeUser();
                if (exchangeUser != null && !string.IsNullOrEmpty(exchangeUser.PrimarySmtpAddress))
                    return exchangeUser.PrimarySmtpAddress;
            }
            catch
            { } // Not an Exchange entry, fall through to the raw address

            return entry.Address ?? "";
        }

        public static bool AddressesMatch(string a, string b)
        {
            return !string.IsNullOrEmpty(a) && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        public static DeliveryType GetDeliveryType(MailItem mailItem)
        {
            string body = mailItem.Body ?? "";

            if (body.Contains("requested link"))
                return DeliveryType.AS_LINK;

            if (body.Contains("requested file attached"))
                return DeliveryType.AS_ATTACHMENT;

            // Body payloads replace the whole email body, so carry neither template marker
            return DeliveryType.AS_BODY;
        }

        public static bool EmailHasAttachment(MailItem mailItem)
        {
            return mailItem.Attachments.Count > 0;
        }

        public static string GetEmailReceivedTime(MailItem mailItem)
        {
            return mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss");
        }

        public static string GetAttachmentHash(MailItem mailItem, string tempPath, bool ondiskFallback = true)
        {
            // Only the first attachment is hashed
            foreach (Attachment attachment in mailItem.Attachments)
            {
                byte[] attachmentData;
                try
                {
                    object property = attachment.PropertyAccessor.GetProperty(PR_ATTACH_DATA_BIN);
                    attachmentData = (byte[])property;
                }
                catch
                {
                    if (!ondiskFallback)
                    {
                        Console.WriteLine("[!] Encountered error calculating hash in memory. Fallback disabled by config. Continuing...");
                        return "";
                    }

                    Console.WriteLine("[!] Encountered error calculating hash in memory. Falling back to on-disk calculation.");

                    // Strip any directory components so the payload is always written inside tempPath
                    string outFile = Path.Combine(tempPath, Path.GetFileName(attachment.FileName));
                    attachment.SaveAsFile(outFile);
                    attachmentData = File.ReadAllBytes(outFile);
                }

                using (MD5 md5 = MD5.Create())
                {
                    byte[] hashBytes = md5.ComputeHash(attachmentData);
                    return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                }
            }
            return "";
        }

        public static string GetAttachmentFileName(MailItem mailItem)
        {
            int attachmentCount = mailItem.Attachments.Count;

            if (attachmentCount > 1)
                Console.WriteLine($"[!] Email has excess attachments. Attachment count is {attachmentCount}.");

            // Only the first attachment is reported
            foreach (Attachment attachment in mailItem.Attachments)
                return attachment.FileName;

            return "";
        }

        public static string GetEmailIdFromSubject(string subject)
        {
            // Subjects contain the campaign ID followed by the email ID
            string[] matches = UuidRegex.Matches(subject ?? "")
                                    .OfType<Match>()
                                    .Select(m => m.Value)
                                    .ToArray();

            if (matches.Length < 2)
            {
                Console.WriteLine("[!] Failed to find email ID in email");
                Console.WriteLine($"[!] Subject: {subject}");
                return "";
            }

            return matches.Last();
        }
    }
}
