using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json.Linq;
using Tomlyn;
using Tomlyn.Model;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace MailCollector
{
    class Program
    {
        // Settings resolved from the config file
        static string senderAddress = "no-reply@delivrto.me";
        static string campaignId;
        static string smtpRecipientAddress = "";
        static string tempPath = "";
        static string outputPath = "";
        static bool cleanup = true;
        static bool msgSave = false;
        static bool deleteEmail = false;
        static bool ondiskFallback = true;

        static readonly object resultsLock = new object();
        static JObject emailResultsJson;
        static JArray emailResultList;

        // Outlook only raises ItemAdd while the Items objects are referenced, so these
        // are held in a static field for the lifetime of monitor mode.
        static Items[] monitoredItems;

        // The results log is written to outputPath, or the current working directory if unset
        private static string LogFilePath()
        {
            return Path.GetFullPath(Path.Combine(outputPath, $"output-{campaignId}.json"));
        }

        private static bool WriteJsonLogToFile()
        {
            string path = LogFilePath();
            try
            {
                string json;
                lock (resultsLock)
                {
                    json = emailResultsJson.ToString();
                }
                File.WriteAllText(path, json + Environment.NewLine);
                return true;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"[!] Error writing \"{path}\": {ex.Message}");
                return false;
            }
        }

        private static JObject ProcessEmail(MailItem mailItem, bool inJunk)
        {
            JObject emailJson = new JObject();

            emailJson["in_junk"] = inJunk;
            emailJson["sent"] = Utils.GetEmailReceivedTime(mailItem);

            // Fetch Email ID from the subject
            string emailId = Utils.GetEmailIdFromSubject(mailItem.Subject);
            if (string.IsNullOrEmpty(emailId))
                return emailJson;

            Console.WriteLine($"[+] {emailId}");
            emailJson["email_id"] = emailId;

            // Identify delivery type from email body
            Utils.DeliveryType mailType = Utils.GetDeliveryType(mailItem);

            string status;

            if (mailType == Utils.DeliveryType.AS_LINK)
            {
                LinkResult linkResult = LinkClassifier.ClassifyLink(mailItem.HTMLBody);
                emailJson["mail_type"] = "as_link";
                emailJson["link"] = linkResult.Link;
                status = linkResult.Status;
            }
            else if (mailType == Utils.DeliveryType.AS_ATTACHMENT)
            {
                emailJson["mail_type"] = "as_attachment";

                if (Utils.EmailHasAttachment(mailItem))
                {
                    // The upload compares the hash to decide between Delivered and Stripped
                    status = EmailStatus.Delivered;

                    string attachmentFileName = Utils.GetAttachmentFileName(mailItem);
                    if (!string.IsNullOrEmpty(attachmentFileName))
                    {
                        string extension = Path.GetExtension(attachmentFileName);
                        emailJson["payload_name"] = attachmentFileName;
                        emailJson["extension"] = extension.Length > 1 ? extension.Substring(1) : "";
                    }

                    string attachmentHash = Utils.GetAttachmentHash(mailItem, tempPath, ondiskFallback);
                    if (!string.IsNullOrEmpty(attachmentHash))
                        emailJson["hash"] = attachmentHash;
                }
                else
                {
                    status = EmailStatus.Held;
                }
            }
            else
            {
                // A body payload that reached the mailbox is reported as delivered
                emailJson["mail_type"] = "as_body";
                status = EmailStatus.Delivered;
            }

            emailJson["status"] = inJunk ? $"Junk ({status})" : status;

            if (msgSave)
            {
                try
                {
                    Utils.SaveMsg(mailItem, tempPath, emailId);
                }
                catch
                {
                    Console.WriteLine($"[!] Failed to save MSG file for email ID: {emailId}");
                }
            }

            return emailJson;
        }

        private static string GetSubject(MailItem mailItem)
        {
            try
            {
                return mailItem.Subject ?? "";
            }
            catch
            {
                return "";
            }
        }

        // Returns true if the mail item belongs to this campaign and was sent to the target mailbox
        private static bool IsCampaignEmail(MailItem mailItem)
        {
            string subject = mailItem.Subject ?? "";
            if (subject.IndexOf(campaignId, StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            AddressEntry sender = mailItem.Sender;
            if (sender == null || !Utils.AddressesMatch(sender.Address, senderAddress))
                return false;

            foreach (Recipient rec in mailItem.Recipients)
            {
                if (Utils.AddressesMatch(Utils.GetSmtpAddress(rec.AddressEntry), smtpRecipientAddress))
                    return true;
            }
            return false;
        }

        // Processes a matched email, records the result and optionally deletes the email
        private static JObject HandleEmail(MailItem mailItem, bool inJunk)
        {
            JObject emailJson;
            try
            {
                emailJson = ProcessEmail(mailItem, inJunk);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"[!] Failed to process email with subject \"{GetSubject(mailItem)}\": {ex.Message}");
                return null;
            }

            lock (resultsLock)
            {
                emailResultList.Add(emailJson);
            }

            // Only delete once the result has been recorded so it is never lost
            if (deleteEmail)
            {
                try
                {
                    mailItem.Delete();
                }
                catch
                {
                    Console.WriteLine($"[!] Failed to delete email with subject \"{GetSubject(mailItem)}\"");
                }
            }

            return emailJson;
        }

        private static void CleanupTempPath()
        {
            if (!cleanup)
                return;

            try
            {
                Directory.Delete(tempPath, true);
            }
            catch
            {
                Console.WriteLine($"[!] Failed to cleanup saved payloads at \"{tempPath}\"");
            }
        }

        private static void WriteLogAndReport()
        {
            if (WriteJsonLogToFile())
                Console.WriteLine($"[+] JSON Log Written to: {LogFilePath()}");
            else
                Console.WriteLine("[!] Failed to write log file!");
        }

        // Returns the string value of a config key, or null if it is absent
        private static string GetConfigString(TomlTable config, string key)
        {
            if (!config.TryGetValue(key, out object value))
                return null;

            if (value is string str)
                return str;

            Console.WriteLine($"[!] \"{key}\" must be a string value. Ignoring.");
            return null;
        }

        // Reads a boolean config key, accepting either a TOML boolean or the string "true"/"false"
        private static bool GetConfigBool(TomlTable config, string key, bool defaultValue)
        {
            if (!config.TryGetValue(key, out object value))
                return defaultValue;

            if (value is bool b)
                return b;

            string str = (value as string ?? "").Trim().ToLowerInvariant();
            if (str == "true")
                return true;
            if (str == "false")
                return false;

            Console.WriteLine($"[!] \"{key}\" must be a value of 'true' or 'false'. Defaulting to '{defaultValue.ToString().ToLowerInvariant()}'.");
            return defaultValue;
        }

        private static MAPIFolder GetDefaultFolder(MAPIFolder root, OlDefaultFolders folderType, string fallbackName)
        {
            // Store.GetDefaultFolder is locale independent; fall back to the English name if unavailable
            try
            {
                return root.Store.GetDefaultFolder(folderType);
            }
            catch
            {
                return root.Folders[fallbackName];
            }
        }

        static void ShowBanner()
        {
            Console.WriteLine(@"

                        ~:
                   ~!777~:^^^:
                !!777777~:^^^^^^^:
            ~!7777777777~:^^^^^^^^^^:
        ~!!7777777777777~:^^^^^^^^^^^^^^:
     :^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^:
     ^~^^^:^^^^^^^^^^^^^^^^^^^^^^^^^^^^::::^^
     ^~~~~^^^^^^^^^^^^^^^^^^^^^^^^^^:::^^^^^^
     ^~~~~~~~~^^^^^^^^^^^^^^^^^^:::^^^^^^^^^^
     ^~~~~~~~~~~~^^^^^^^^^^^^:::^^^^^^^^^^^^^     MailCollector
     ^~~~~~~~~~~~~~~~^^^^:::^^^^^^^^^^^^^^^^^
     ^~~~~~~~~~~~~~~~~~~^:^^^^^^^^^^^^^^^^^^^     delivr.to
     ^~~~~~~~~~~~~~~~~~~^:^^^^^^^^^^^^^^^^^^^
     ^~~~~~~~~~~~~~~~~~~^:^^^^^^^^^^^^^^^^^^^
     ^~~~~~~~~~~~~~~~~~~^:^^^^^^^^^^^^^^^^^^^
     ~~~~~~~~~~~~~~~~~~~^:^^^^^^^^^^^^^^^^^^^
      ^~~~~~~~~~~~~~~~~~^:^^^^^^^^^^^^^^^^^
         ^^~~~~~~~~~~~~~^:^^^^^^^^^^^^^^
             ^~~~~~~~~~~^:^^^^^^^^^^
                ^^~~~~~~^:^^^^^^^
                    ^~~~^:^^^
                         ^

");
        }

        static void Main(string[] args)
        {
            const string usage = "MailCollector.exe <capture|monitor> [config.toml]";
            string mode;
            string configFile = "config.toml";

            ShowBanner();

            if (args.Length < 1 || args.Length > 2)
            {
                Console.WriteLine($"[!] Usage: {usage}");
                return;
            }

            if (args[0] == "monitor")
            {
                Console.WriteLine("[+] Monitor mode selected. Incoming mail will be output to log.");
                mode = "monitor";
            }
            else if (args[0] == "capture")
            {
                Console.WriteLine("[+] Capture mode selected. Existing mail will be captured for processing.");
                mode = "capture";
            }
            else
            {
                Console.WriteLine($"[!] Incorrect mode requested. Usage: {usage}");
                return;
            }

            if (args.Length == 2)
                configFile = args[1];

            // Load configuration from TOML file
            TomlTable config;
            try
            {
                config = Toml.ToModel(File.ReadAllText(configFile), configFile);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"[!] Failed to load config file \"{configFile}\": {ex.Message}");
                return;
            }

            string configSender = GetConfigString(config, "senderAddress");
            if (string.IsNullOrEmpty(configSender))
                Console.WriteLine($"[+] No sending email in config, defaulting to \"{senderAddress}\".");
            else
                senderAddress = configSender;

            // Retrieve and validate the campaign ID
            string configCampaignId = (GetConfigString(config, "campaignId") ?? "").Trim();
            if (string.IsNullOrEmpty(configCampaignId))
            {
                Console.WriteLine("[!] Failed to find campaign ID in config file. Exiting...");
                return;
            }
            if (configCampaignId.Length > 36)
                configCampaignId = configCampaignId.Substring(0, 36);
            if (!Guid.TryParseExact(configCampaignId, "D", out _))
            {
                Console.WriteLine("[!] Invalid Campaign ID UUID provided. Exiting...");
                return;
            }
            campaignId = configCampaignId.ToLowerInvariant();
            Console.WriteLine($"[+] Searching for emails with campaign ID: {campaignId}.");

            string folderName = GetConfigString(config, "folderName");
            if (string.IsNullOrEmpty(folderName))
                Console.WriteLine("[+] No folder name in config, defaulting to \"Inbox\".");

            NameSpace outlookNamespace;
            try
            {
                Application outlookApp = new Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
            }
            catch
            {
                Console.WriteLine("[!] Failed to access Outlook. Exiting...");
                return;
            }

            // Resolve the mailbox to search
            string recipientAddress = GetConfigString(config, "recipientAddress");
            try
            {
                if (string.IsNullOrEmpty(recipientAddress))
                {
                    smtpRecipientAddress = Utils.GetSmtpAddress(outlookNamespace.CurrentUser.AddressEntry);
                    Console.WriteLine($"[+] No recipient email in config, using default \"{smtpRecipientAddress}\".");
                }
                else
                {
                    foreach (Account account in outlookNamespace.Accounts)
                    {
                        if (Utils.AddressesMatch(account.SmtpAddress, recipientAddress))
                        {
                            smtpRecipientAddress = account.SmtpAddress;
                            Console.WriteLine($"[+] Found account: {account.SmtpAddress}");
                            break;
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine($"[!] Failed to access mailbox for \"{recipientAddress}\". Exiting...");
                return;
            }

            if (string.IsNullOrEmpty(smtpRecipientAddress))
            {
                Console.WriteLine($"[!] No Outlook account found for \"{recipientAddress}\". Exiting...");
                return;
            }

            MAPIFolder targetFolder;
            MAPIFolder junkFolder = null;
            try
            {
                MAPIFolder mailboxRoot = outlookNamespace.Folders[smtpRecipientAddress];
                MAPIFolder inbox = GetDefaultFolder(mailboxRoot, OlDefaultFolders.olFolderInbox, "Inbox");

                if (string.IsNullOrEmpty(folderName))
                {
                    targetFolder = inbox;
                }
                else
                {
                    // Try both inside and outside of 'Inbox' before giving up on locating folder
                    try
                    {
                        targetFolder = inbox.Folders[folderName];
                    }
                    catch
                    {
                        targetFolder = mailboxRoot.Folders[folderName];
                    }
                }

                try
                {
                    junkFolder = GetDefaultFolder(mailboxRoot, OlDefaultFolders.olFolderJunk, "Junk Email");
                }
                catch
                {
                    Console.WriteLine("[!] Failed to find the Junk Email folder, only the target folder will be searched.");
                }
            }
            catch
            {
                Console.WriteLine("[!] Failed to find target folder, does it exist?");
                return;
            }

            emailResultsJson = new JObject(
                new JProperty("campaign_id", campaignId),
                new JProperty("results", new JArray())
            );
            emailResultList = (JArray)emailResultsJson["results"];

            // Resolve where the results log is written
            string outputPathStr = GetConfigString(config, "outputPath");
            if (!string.IsNullOrEmpty(outputPathStr))
            {
                try
                {
                    outputPath = Path.GetFullPath(outputPathStr);
                    Directory.CreateDirectory(outputPath);
                    Console.WriteLine($"[+] Using specified results output path at \"{outputPath}\"");
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine($"[!] Failed to use the specified results output path ({ex.Message}), falling back to the current directory.");
                    outputPath = "";
                }
            }

            // Resolve where payloads are saved
            try
            {
                string tempPathStr = GetConfigString(config, "tempPath");
                if (string.IsNullOrEmpty(tempPathStr))
                {
                    tempPath = Utils.CreateUniqueTempDirectory();
                }
                else
                {
                    tempPath = Utils.CreateUniqueTempDirectory(Utils.NormalisePath(tempPathStr));
                    Console.WriteLine($"[+] Using specified payload save path at \"{tempPath}\"");
                }
            }
            catch
            {
                Console.WriteLine("[!] Failed to use the specified payload save path, falling back to %TEMP%.");
                tempPath = Utils.CreateUniqueTempDirectory();
            }

            // Allow for fallback to be blocked so no attachment is written to disk
            ondiskFallback = GetConfigBool(config, "ondiskFallback", true);
            if (!ondiskFallback)
                Console.WriteLine("[+] No fallback to on-disk attachment hash calculation.");

            // Allow for cleanup to be skipped so attachments are retained
            cleanup = GetConfigBool(config, "cleanup", true);
            if (!cleanup)
                Console.WriteLine("[+] Cleanup to be skipped.");

            // Save messages as part of processing
            if (GetConfigBool(config, "msgSave", false))
            {
                if (cleanup)
                {
                    Console.WriteLine("[!] MSG files set to save, but cleanup is also true. One of these must be false, reverting MSG save to false.");
                }
                else
                {
                    Console.WriteLine("[+] MSG files will be saved to output directory.");
                    msgSave = true;
                }
            }

            // Allow for email to be deleted if specified
            deleteEmail = GetConfigBool(config, "deleteEmail", false);
            if (deleteEmail)
                Console.WriteLine("[+] Emails to be deleted once parsed.");

            // Junk is searched alongside the target folder; the flag marks results found there
            var folders = new List<KeyValuePair<MAPIFolder, bool>>();
            if (junkFolder != null)
                folders.Add(new KeyValuePair<MAPIFolder, bool>(junkFolder, true));
            folders.Add(new KeyValuePair<MAPIFolder, bool>(targetFolder, false));

            Console.WriteLine($"\n[+] Searching '{targetFolder.Name}' folder.\n");

            if (mode == "capture")
            {
                RunCapture(folders);
            }
            else
            {
                RunMonitor(folders);
            }
        }

        private static void RunCapture(List<KeyValuePair<MAPIFolder, bool>> folders)
        {
            // Collect matches before processing: deleting items while enumerating an
            // Outlook Items collection causes the enumerator to skip items.
            var matches = new List<KeyValuePair<MailItem, bool>>();
            foreach (var folder in folders)
            {
                foreach (object obj in folder.Key.Items)
                {
                    MailItem mailItem = obj as MailItem;
                    if (mailItem == null)
                        continue;

                    try
                    {
                        if (IsCampaignEmail(mailItem))
                            matches.Add(new KeyValuePair<MailItem, bool>(mailItem, folder.Value));
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine($"[!] Failed to inspect email with subject \"{GetSubject(mailItem)}\": {ex.Message}");
                    }
                }
            }

            foreach (var match in matches)
                HandleEmail(match.Key, match.Value);

            Console.WriteLine($"\n[+] Processed {emailResultList.Count} delivr.to emails.");

            CleanupTempPath();
            WriteLogAndReport();
        }

        private static void RunMonitor(List<KeyValuePair<MAPIFolder, bool>> folders)
        {
            Console.CancelKeyPress += delegate
            {
                CleanupTempPath();
                WriteLogAndReport();
            };

            monitoredItems = new Items[folders.Count];
            for (int i = 0; i < folders.Count; i++)
            {
                bool inJunk = folders[i].Value;
                monitoredItems[i] = folders[i].Key.Items;
                monitoredItems[i].ItemAdd += (object item) => OnItemAdded(item, inJunk);
            }

            Console.WriteLine("[+] Monitoring inbox for incoming emails. Press Ctrl+C to stop.");
            Thread.Sleep(Timeout.Infinite);
        }

        private static void OnItemAdded(object item, bool inJunk)
        {
            // Exceptions must not escape an Outlook event handler or the process terminates
            try
            {
                MailItem mailItem = item as MailItem;
                if (mailItem == null || !IsCampaignEmail(mailItem))
                    return;

                JObject emailJson = HandleEmail(mailItem, inJunk);
                if (emailJson == null)
                    return;

                Console.WriteLine(emailJson.ToString());

                // Persist after every result so nothing is lost if the window is closed
                WriteJsonLogToFile();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"[!] Failed to handle incoming item: {ex.Message}");
            }
        }
    }
}
