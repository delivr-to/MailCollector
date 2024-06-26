﻿using System;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using Nett;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Newtonsoft.Json.Linq;
using static System.Guid;
using System.Security.Cryptography;

namespace MailCollector
{
    class Program
    {
        private static bool WriteJsonLogToFile(JObject results, string campaignId)
        {
            string path = $"output-{campaignId}.json";
            try
            {
                using (StreamWriter sw = new StreamWriter(new FileStream(path, FileMode.Create, FileAccess.Write)))
                {
                    sw.WriteLine(results.ToString());
                }
                return true;
            }
            catch
            {
                return true;            
            }
        }

        private static JObject ProcessEmail(MailItem mailItem, string tempPath, bool ondiskFallback=true)
        {
            dynamic emailJson = new JObject();

            if (mailItem.Parent().ToString().Contains("Junk"))
                emailJson.in_junk = true;
            else
                emailJson.in_junk = false;

            emailJson.sent = Utils.GetEmailReceivedTime(mailItem);

            // Fetch Email ID from the subject
            string emailId = Utils.GetEmailIdFromSubject(mailItem.Subject);
            if (!string.IsNullOrEmpty(emailId)){
                Console.WriteLine($"[+] {emailId}");
                emailJson.email_id = emailId;
            } else
            {
                return emailJson;
            }

            // Identify delivery type from email body
            Utils.DeliveryType mailType = Utils.GetDeliveryType(mailItem);
            if (mailType != Utils.DeliveryType.UNKNOWN)
            {
                if(mailType == Utils.DeliveryType.AS_LINK)
                {
                    // Email is a link
                    emailJson.mail_type = "as_link";
                    emailJson.link = Utils.GetLinkFromEmail(mailItem);
                }
                else if (mailType == Utils.DeliveryType.AS_ATTACHMENT)
                {
                    // Email is an attachment
                    emailJson.mail_type = "as_attachment";
                }
            }
            else
            {
                emailJson.status = "Held";

                // Unable to identify either link or attachment type from email body
                return emailJson;
            }

            if (mailType == Utils.DeliveryType.AS_LINK)
            {
                Utils.LinkDeliveryState emailLink = Utils.GetEmailedLinkState(mailItem);
                switch (emailLink)
                {
                    case Utils.LinkDeliveryState.DELIVERED:
                        emailJson.status = "Delivered";
                        break;
                    case Utils.LinkDeliveryState.REWRITTEN:
                        emailJson.status = "Rewritten";
                        break;
                    case Utils.LinkDeliveryState.HELD:
                        emailJson.status = "Held";
                        break;
                    default:
                        break;
                };
            } 
            else if (mailType == Utils.DeliveryType.AS_ATTACHMENT)
            {
                bool hasAttachment = Utils.EmailHasAttachment(mailItem);
                if (hasAttachment)
                {
                    string attachmentFileName = Utils.GetAttachmentFileName(mailItem);
                    if (!string.IsNullOrEmpty(attachmentFileName))
                    {
                        emailJson.payload_name = attachmentFileName;
                        emailJson.extension = Path.GetExtension(attachmentFileName).Substring(1);
                        emailJson.status = "Delivered";
                    }

                    string attachmentHash = Utils.GetAttachmentHash(mailItem, tempPath, ondiskFallback);
                    if (!string.IsNullOrEmpty(attachmentHash))
                    {
                        emailJson.hash = attachmentHash;
                    }
                } else
                {
                    emailJson.status = "Held";
                }
            }
            if ((bool)emailJson.in_junk)
                emailJson.status = $"Junk ({emailJson.status})";
            return emailJson;
        }

        static void showBanner()
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
            string mode = "capture";
            string senderAddress = "no-reply@delivrto.me";
            string recipientAddress = "";
            string campaignId;
            string folderName = "";
            string configFile = "config.toml";
            string smtpRecipientAddress = "";
            bool cleanup = true;
            bool deleteEmail = false;
            bool ondiskFallback = true;
            string tempPath = "";

            showBanner();

            if (args.Length >= 3 || args.Length < 1)
            {
                Console.WriteLine("[-] Missing config file. MailCollector.exe <mode> <config.toml>");
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
            }
            else { 
                Console.WriteLine("[!] Incorrect mode requested. MailCollector.exe <mode> <config.toml>");
                return;
            }

            // Load configuration from TOML file
            TomlTable config;

            try
            {
                if (args.Length == 2)
                {
                    config = Toml.ReadFile(args[1]);
                }
                else
                {
                    config = Toml.ReadFile(configFile);
                }
                
            } catch
            {
                Console.WriteLine("[+] Failed to load config file. Exiting...");
                return;
            }
                
            
            try
            {
                // Retrieve sender address from configuration
                senderAddress = config.Get<string>("senderAddress");
            }
            catch
            {
                Console.WriteLine($"[+] No sending email in config, defaulting to \"{senderAddress}\".");
            }

            
            try
            {
                // Retrieve campaign ID from configuration
                campaignId = config.Get<string>("campaignId").Substring(0, 36);
                
                // Confirm valid UUID before proceeding
                if(!Guid.TryParse(campaignId, out Guid output))
                {
                    Console.WriteLine("[!] Invalid Campaign ID UUID provided. Exiting...");
                    return;
                }

                Console.WriteLine($"[+] Searching for emails with campaign ID: {campaignId}.");
            }
            catch
            {
                Console.WriteLine("[!] Failed to find campaign ID in config file. Exiting...");
                return;
            }

            try
            {
                // Retrieve folder name from configuration
                folderName = config.Get<string>("folderName");
            }
            catch
            {
                Console.WriteLine("[+] No folder name in config, defaulting to \"Inbox\".");
            }

            
            Application outlookApp;
            NameSpace outlookNamespace;
            MAPIFolder targetFolder;
            MAPIFolder junkFolder;
            Items items;
            Items junkItems;

            try
            {
                outlookApp = new Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
            }
            catch
            {
                Console.WriteLine($"[!] Failed to access Outlook. Exiting...");
                return;
            }

         
            try
            {
                // Retrieve recipient address from configuration
                recipientAddress = config.Get<string>("recipientAddress");
            }
            catch
            {
                smtpRecipientAddress = outlookNamespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                Console.WriteLine($"[+] No recipient email in config, using default \"{smtpRecipientAddress}\".");
            }

            try
            {
                // Allow for fallback to be blocked so no attachment is written to disk.
                string ondiskFallbackStr = config.Get<string>("ondiskFallback").ToLower();

                if (ondiskFallbackStr == "false")
                {
                    Console.WriteLine($"[+] No fallback to on-disk attachment hash calculation.");
                    ondiskFallback = false;
                }
                else if (ondiskFallbackStr != "true")
                {
                    Console.WriteLine($"[!] \"ondiskFallbackStr\" must be a value of 'true' or 'false'. Defaulting to 'true'.");
                }
            }
            catch
            { } // Will fallback as usual

            try
            {
                if (String.IsNullOrEmpty(smtpRecipientAddress))
                {
                    foreach (Account user in outlookNamespace.Accounts)
                    {
                        if (user.SmtpAddress.ToLower() == recipientAddress.ToLower())
                        {
                            smtpRecipientAddress = user.SmtpAddress;
                            Console.WriteLine($"[+] Found account: {user.SmtpAddress}");
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine($"[+] Failed to access mailbox for \"{recipientAddress}\". Exiting...");
                return;
            }

            try
            {
                if (string.IsNullOrEmpty(folderName))
                {
                    targetFolder = outlookNamespace.Folders[smtpRecipientAddress].Folders["Inbox"];
                }
                else
                {
                    //Try both inside and outside of 'Inbox' before giving up on locating folder
                    try
                    {
                        targetFolder = outlookNamespace.Folders[smtpRecipientAddress].Folders["Inbox"].Folders[folderName];
                    }
                    catch
                    {
                        targetFolder = outlookNamespace.Folders[smtpRecipientAddress].Folders[folderName];
                    }
                    
                }
            }
            catch
            {
                Console.WriteLine($"[!] Failed to find target folder, does it exist?");
                return;
            }

            junkFolder = outlookNamespace.Folders[smtpRecipientAddress].Folders["Junk Email"];
            
            junkItems = junkFolder.Items;
            items = targetFolder.Items;
            Items[] itemSets = { junkItems, items };

            JObject emailResultsJson = new JObject(
                new JProperty("campaign_id", campaignId),
                new JProperty("results", new JArray())
            );

            JArray emailResultList = emailResultsJson["results"].Value<JArray>();

            try
            {
                string tempPathStr = config.Get<string>("tempPath");
                if (String.IsNullOrEmpty(tempPathStr))
                {
                    tempPath = Utils.CreateUniqueTempDirectory();
                } else
                {
                    string NormPath = Utils.NormalisePath(tempPathStr);
                    tempPath = Utils.CreateUniqueTempDirectory(NormPath);
                    Console.WriteLine($"[+] Using specified payload save path at \"{tempPath}\"");
                }
            } catch
            {
                tempPath = Utils.CreateUniqueTempDirectory();
            }
            

            try
            {
                // Allow for cleanup to be skipped so attachments are retained.
                string cleanupStr = config.Get<string>("cleanup").ToLower();
                
                if (cleanupStr == "false")
                {
                    Console.WriteLine($"[+] Cleanup to be skipped.");
                    cleanup = false;
                }
                else if (cleanupStr != "true")
                {
                    Console.WriteLine($"[!] \"cleanup\" must be a value of 'true' or 'false'. Defaulting to 'true'.");
                }
            }
            catch
            { } // Will cleanup as usual

            try
            {
                // Allow for email to be deleted if specified.
                string deleteStr = config.Get<string>("deleteEmail").ToLower();

                if (deleteStr == "true")
                {
                    Console.WriteLine($"[+] Emails to be deleted once parsed.");
                    deleteEmail = true;
                }
                else if (deleteStr != "false")
                {
                    Console.WriteLine($"[!] \"deleteEmail\" must be a value of 'true' or 'false'. Defaulting to 'false'.");
                }
            }
            catch
            { } // Will skip email deletion by default

            Console.WriteLine($"\n[+] Searching '{targetFolder.Name}' folder.\n");

            if (mode == "capture")
            {
                foreach (Items itemSet in itemSets)
                {
                    foreach (Object _obj in itemSet)
                    {
                        if (_obj is MailItem)
                        {
                            MailItem mailItem = (MailItem) _obj;
                            if (mailItem.Sender.Address == senderAddress && mailItem.Subject.Contains(campaignId))
                            {
                                foreach(Recipient rec in mailItem.Recipients)
                                {
                                    if(rec.AddressEntry.GetExchangeUser().PrimarySmtpAddress == smtpRecipientAddress)
                                    {
                                        try
                                        {
                                            JObject emailJson = ProcessEmail(mailItem, tempPath, ondiskFallback);
                                            emailResultList.Add(emailJson);

                                            // Won't delete email if there's a prior processing issue as we'd lose the result
                                            if (deleteEmail)
                                            {
                                                try
                                                {
                                                    mailItem.Delete();
                                                } catch
                                                {
                                                    Console.WriteLine($"Failed to delete email with subject \"{mailItem.Subject}\"");
                                                }
                                            }
                                        }
                                        catch
                                        {
                                            Console.WriteLine($"Failed to process email with subject \"{mailItem.Subject}\"");
                                        }
                                        
                                    }
                                }
                            }
                        }
                    }
                }
                Console.WriteLine($"\n[+] Processed {emailResultList.Count} delivr.to emails.");

                if(cleanup)
                {
                    try
                    {
                        Directory.Delete(tempPath, true);
                    }
                    catch
                    {
                        Console.WriteLine($"[!] Failed to cleanup saved payloads at \"{tempPath}\"");
                    }
                }

                bool logWritten = WriteJsonLogToFile(emailResultsJson, campaignId);

                if (logWritten)
                    Console.WriteLine($"[+] JSON Log Written to: output-{campaignId}.json");
                else
                    Console.WriteLine("[!] Failed to write log file!");
            }
            else
            {
                Console.CancelKeyPress += delegate {
                    if (cleanup)
                    {
                        try
                        {
                            Directory.Delete(tempPath, true);
                        }
                        catch
                        {
                            Console.WriteLine($"[!] Failed to cleanup saved payloads at \"{tempPath}\"");
                        }
                    }

                    bool logWritten = WriteJsonLogToFile(emailResultsJson, campaignId);
                    if (logWritten)
                        Console.WriteLine($"[+] JSON Log Written to: output-{campaignId}.json");
                    else
                        Console.WriteLine("[!] Failed to write log file!");
                };

                foreach (Items itemSet in itemSets)
                {
                    itemSet.ItemAdd += (object item) =>
                    {
                        MailItem mailItem = (MailItem)item;
                        if (mailItem.Sender.Address == senderAddress && mailItem.Subject.Contains(campaignId))
                        {
                            foreach (Recipient rec in mailItem.Recipients)
                            {
                                if (rec.AddressEntry.GetExchangeUser().PrimarySmtpAddress == smtpRecipientAddress)
                                {
                                    try
                                    {
                                        JObject emailJson = ProcessEmail(mailItem, tempPath, ondiskFallback);
                                        emailResultList.Add(emailJson);
                                        Console.WriteLine(emailJson.ToString());

                                        // Won't delete email if there's a prior processing issue as we'd lose the result
                                        if (deleteEmail)
                                        {
                                            try
                                            {
                                                mailItem.Delete();
                                            }
                                            catch
                                            {
                                                Console.WriteLine($"Failed to delete email with subject \"{mailItem.Subject}\"");
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        Console.WriteLine($"Failed to process email with subject \"{mailItem.Subject}\"");
                                    }
                                }
                            }
                        }
                    };
                };
                Console.WriteLine("[+] Monitoring inbox for incoming emails...");
                while (true)
                {
                    Console.ReadLine();
                }
            }
        }
    }
}