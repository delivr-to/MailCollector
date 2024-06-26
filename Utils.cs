﻿
using System;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using System.Security.Cryptography;
using System.IO;


namespace MailCollector
{

    class Utils
    {
        static string LINK_DOMAIN = "delivrto\\.me";

        //https://stackoverflow.com/questions/29912136/converting-a-outlook-mail-attachment-to-byte-array-with-c-sharp
        const string PR_ATTACH_DATA_BIN = "http://schemas.microsoft.com/mapi/proptag/0x37010102";

        public enum DeliveryType
        {
            AS_LINK,
            AS_ATTACHMENT,
            UNKNOWN
        }

        public enum LinkDeliveryState
        {
            DELIVERED,
            REWRITTEN,
            HELD
        }

        public static string NormalisePath(string tempPath)
        {
            var path = Path.GetFullPath(tempPath);
            if (!Directory.Exists(path))
            {
                Console.WriteLine($"[+] Payload save directory doesn't exist, attempting to create it.");
                Directory.CreateDirectory(tempPath);
            }
            return path;
        }

        public static string CreateUniqueTempDirectory(string path="")
        {
            string uniqueTempDir;

            if (String.IsNullOrEmpty(path))
                uniqueTempDir = Path.GetFullPath(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));
            else
                uniqueTempDir = Path.Combine(path, Guid.NewGuid().ToString());
            
            Directory.CreateDirectory(uniqueTempDir);
            return uniqueTempDir;
        }

        public static DeliveryType GetDeliveryType(MailItem mailItem)
        {
            if(mailItem.Body.Contains("requested link"))
            {
                return DeliveryType.AS_LINK;
            } 
            else if (mailItem.Body.Contains("requested file attached"))
            {
                return DeliveryType.AS_ATTACHMENT;
            }
            else 
            {
                return DeliveryType.UNKNOWN;
            }
        }

        public static bool EmailHasAttachment(MailItem mailItem)
        {
            return mailItem.Attachments.Count > 0;
        }

        public static string GetEmailReceivedTime(MailItem mailItem)
        {
            return mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss");
        }

        public static string GetAttachmentHash(MailItem mailItem, string tempPath, bool ondiskFallback=true)
        {
            // Return after the first one
            foreach (Attachment attachment in mailItem.Attachments)
            {
                byte[] attachmentData = null;
                try
                {
                    attachmentData = attachment.PropertyAccessor.GetProperty(PR_ATTACH_DATA_BIN);
                } catch
                {
                    if (ondiskFallback)
                    {
                        Console.WriteLine($"[!] Encountered error calculating hash in memory. Falling back to on-disk calculation.");
                        string outFile = Path.Combine(tempPath, attachment.FileName);
                        attachment.SaveAsFile(outFile);
                        attachmentData = File.ReadAllBytes(outFile);
                    } else
                    {
                        Console.WriteLine($"[!] Encountered error calculating hash in memory. Fallback disabled by config. Continuing...");
                        return "";
                    }
                }

                using (var md5 = MD5.Create())
                {
                    
                    byte[] hashBytes = md5.ComputeHash(attachmentData);
                    string hashString = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                    return hashString;
                }
            }
            return "";
        }

        public static string GetAttachmentFileName(MailItem mailItem)
        {
            int attachmentCount = mailItem.Attachments.Count;
            
            if (attachmentCount > 1)
            {
                Console.WriteLine($"[!] Email has excess attachments. Attachment count is {attachmentCount}.");
            }

            // Return after the first one
            foreach (Attachment attachment in mailItem.Attachments)
            {
                return attachment.FileName;
            }
            return "";
        }

        public static string GetLinkFromEmail(MailItem mailItem)
        {
            Regex link_regex = new Regex("https:\\/\\/"+ LINK_DOMAIN+ "\\/links\\/[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}");
            
            string[] link_matches = link_regex.Matches(mailItem.HTMLBody)
                                    .OfType<Match>()
                                    .Select(m => m.Groups[0].Value)
                                    .ToArray();
            if (link_matches.Length >= 1)
                return link_matches.First();
            
            return "";
        }

        public static LinkDeliveryState GetEmailedLinkState(MailItem mailItem)
        {
            Regex link_regex = new Regex("https:\\/\\/" + LINK_DOMAIN + "\\/links\\/[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}");
            Regex valid_link_html = new Regex("<a\\s*href=\\\"https:\\/\\/" + LINK_DOMAIN + "\\/links\\/[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}\\\">https:\\/\\/" + LINK_DOMAIN + "\\/links\\/[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}<\\/a>");
            

            string linkOutput = "";
            
            string[] valid_link_matches = valid_link_html.Matches(mailItem.HTMLBody)
                                    .OfType<Match>()
                                    .Select(m => m.Groups[0].Value)
                                    .ToArray();

            if(valid_link_matches.Length == 1)
            {
                linkOutput = valid_link_matches.First();
                return Utils.LinkDeliveryState.DELIVERED;
            }

            string[] link_matches = link_regex.Matches(mailItem.HTMLBody)
                                    .OfType<Match>()
                                    .Select(m => m.Groups[0].Value)
                                    .ToArray();
            if (link_matches.Length == 1)
            {
                linkOutput = link_matches.First();
                return Utils.LinkDeliveryState.REWRITTEN;
            }

            return Utils.LinkDeliveryState.HELD;
        }

        public static string GetEmailIdFromSubject(string subject)
        {
            Regex validateUUIDRegex = new Regex("^[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}$");
            
            // Extract UUID from a string
            Regex extractUUIDRegex = new Regex("[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}");
            string[] matches = extractUUIDRegex.Matches(subject)
                                    .OfType<Match>()
                                    .Select(m => m.Groups[0].Value)
                                    .ToArray();

            if (matches.Length < 2)
            {
                Console.WriteLine("[!] Failed to find email ID in email");
                Console.WriteLine($"[!] Subject: {subject}");
                return "";
            }

            string emailId = matches.Last().ToString();
            return emailId;
        }
    }
}
