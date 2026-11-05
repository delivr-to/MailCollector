# MailCollector

This .NET 4.8 application allows for the 'offline' collection of delivr.to campaign results from a Windows endpoint running Microsoft Outlook.

MailCollector leverages COM to communicate with the Outlook client and retrieve results. It can triage links, attachments and body payloads, and supports filtering for results by campaign ID, sender, recipient, and mailbox folder.

> Read more about MailCollector in the release blog, [here](https://blog.delivr.to/mailcollector-collecting-delivr-to-results-from-outlook-7a28630dcd69).

## Installation

Prebuilt releases can be downloaded from the [Releases](https://github.com/delivr-to/MailCollector/releases) page. Each release zip contains `MailCollector.exe`, its `.exe.config` and an example config.

To build from source: MailCollector targets .NET Framework 4.8 and builds with the [.NET SDK](https://dotnet.microsoft.com/download) (8.0 or later) or Visual Studio 2022.

1. Clone the repository.
2. In the project folder run `dotnet build -c Release` (alternatively, open `MailCollector.sln` and build the `Release` configuration in Visual Studio).
3. Optionally, run the unit tests with `dotnet test`.

The resulting `bin\Release\MailCollector.exe` is a single self-contained executable, with its dependencies embedded by [Costura](https://github.com/Fody/Costura). It requires Windows, the .NET Framework 4.8 runtime and Microsoft Outlook to run.

## Usage

### Configuration

MailCollector uses a TOML file (`config.toml`) to tailor configuration without requiring the tool to be recompiled with hardcoded values. Unless specified otherwise, MailCollector will use the `config.toml` file in the same directory. This configuration file can contain the following values:

```toml
campaignId = "d373b79e-bd9a-4e3a-89e7-a14bdd3df9e3"
senderAddress = "no-reply@delivrto.me"
recipientAddress = "test-mailbox@delivr.to"
folderName = "delivr.to"
cleanup = "true"
tempPath = "payloads"
outputPath = "output"
deleteEmail = "false"
ondiskFallback = "true"
msgSave = "false"
```

At a minimum, the configuration file needs to contain the `campaignId` value. Which can be retrieved from delivr.to either retrospectively, or by scheduling a future campaign (relevant for 'monitor' mode, see below).

The other options are optional and achieve the following. Boolean options accept either a TOML boolean (`true`) or a string (`"true"`):

- `senderAddress` - Defaults to `no-reply@delivrto.me`, but can be changed for custom senders.
- `recipientAddress` - Defaults to the primary Outlook mailbox, but useful if you have more than one mailbox configured in Outlook.
- `folderName` - Defaults to the inbox, but can be filtered to a specific folder, e.g. if rules have been configured.
- `cleanup` - Defaults to "true", but can be configured to "false" to retain saved payloads used in hash calculation, e.g. for debugging purposes or further analysis.
- `tempPath` - Where payloads are saved while being processed. Defaults to a GUID-named folder in %TEMP%, but a relative or absolute path can be given to customise this (e.g. for AV exclusions). A GUID-named subfolder is created inside it for each run. Backslashes must be escaped.
- `outputPath` - Where the `output-[CAMPAIGN_ID].json` results file is written. Defaults to the current working directory. The folder is created if it doesn't exist. Backslashes must be escaped.
- `deleteEmail` - Defaults to "false", but can be configured to delete each delivr.to email once processed.
- `ondiskFallback` - Defaults to "true", attachment hash calculation will be attempted in-memory, any issues with this will fallback to writing to disk unless this value is set to "false".
- `msgSave` - Defaults to "false" and requires `cleanup` to be false. This saves the MSG file for each email to the run's subfolder of `tempPath`.

### Execution

MailCollector operates in two similar but distinct modes, `capture` and `monitor`.

In `capture` mode, MailCollector can be run retrospectively, after a campaign has completed, to retrieve results. In `monitor` mode, arriving emails are actively monitored and analysed. Results are written to an `output-[CAMPAIGN_ID].json` log file, after each new result in `monitor` mode and once processing completes in `capture` mode. Press Ctrl+C to stop `monitor` mode.

`capture` mode execution:

```
> .\MailCollector.exe capture [config.toml] 
```

`monitor` mode execution:

```
> .\MailCollector.exe monitor [config.toml] 
```

An example of `capture` mode execution can be seen below:

```
> .\MailCollector.exe capture


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


[+] Capture mode selected. Existing mail will be captured for processing.
[+] Searching for emails with campaign ID: d373b79e-bd9a-4e3a-89e7-a14bdd3df9e3.
[+] Found account: test-mailbox@delivr.to

[+] Searching 'delivr.to' folder.
[+] 00cee906-ef9d-4e11-bf02-95a51e6ab845
[+] 327613cd-f859-4d15-bd12-5e0d2c50e424
...
[+] 1b25741f-8426-4bcd-ba20-cfad2de8f90d

[+] Processed 15 delivr.to emails.
[+] JSON Log Written to: output-d373b79e-bd9a-4e3a-89e7-a14bdd3df9e3.json
```

> NOTE: For attachment analysis, MailCollector saves payloads to a new directory in %TEMP%, by default. This can be customised with the `tempPath` config entry. Given AV/EDR may quarantine files that were otherwise permitted by mail gateways, it is recommended that appropriate allowlisting is in place to prevent result processing issues.

### Output

Upon results processing completion an `output-[CAMPAIGN_ID].json` file is created in `outputPath`, or the current working directory if that isn't set. An example of its output can be seen below:

```json
{ 
    "campaign_id" : "d373b79e-bd9a-4e3a-89e7-a14bdd3df9e3",
    "results": [
        {
            "in_junk": false,
            "sent": "2023-05-24 19:54:50",
            "email_id": "6e8bb035-c50c-47c8-b56c-48332805b666",
            "mail_type": "as_link",
            "link": "https://delivrto.me/links/a6882f3b-d165-469a-987f-4a5f7a94f2aa",
            "status": "Delivered"
        },
        {
            "in_junk": false,
            "sent": "2023-05-24 19:54:49",
            "email_id": "a17e6ee1-0b06-4fbc-bf1b-ff4b323ef69d",
            "mail_type": "as_attachment",
            "payload_name": "Benefit.html",
            "extension": "html",
            "status": "Delivered",
            "hash": "f3f96be1a14f905195dbc0c1c610ddff"
        },
        {
            "in_junk": false,
            "sent": "2023-05-24 19:54:51",
            "email_id": "0b1e3b0c-6a43-4d5e-9f3e-2f6c1f0a7d21",
            "mail_type": "as_body",
            "status": "Delivered"
        }
    ]
}
```

Each result's `status` is determined as follows, matching how the delivr.to Microsoft 365 and Google integrations classify results:

- **Links** - `Delivered` if the link still opens the original URL, `Rewritten` if the original URL is present but the link was rewritten (e.g. by Safe Links), otherwise `Held`. For rewritten links, `link` contains the rewritten URL.
- **Attachments** - `Delivered` if an attachment is present, otherwise `Held`. The `hash` is compared by delivr.to on upload, which marks the result `Stripped` if the attachment was modified.
- **Body payloads** - `Delivered` when the email is found.

Results found in the Junk Email folder are reported as `Junk (<status>)`.

This `output-[CAMPAIGN_ID].json` can be uploaded directly in the delivr.to `Campaign Details` UI in order to synchronise results.

Notably several fields in this JSON are provided for utility to end users, and aren't required for processing in the UI. A minimal results JSON is as follows:

```json
{ 
    "campaign_id" : "d373b79e-bd9a-4e3a-89e7-a14bdd3df9e3",
    "results": [
        {
            "email_id": "6e8bb035-c50c-47c8-b56c-48332805b666",
            "link": "https://delivrto.me/links/a6882f3b-d165-469a-987f-4a5f7a94f2aa"
        },
        {
            "email_id": "a17e6ee1-0b06-4fbc-bf1b-ff4b323ef69d",
            "hash": "f3f96be1a14f905195dbc0c1c610ddff"
        },
        {
            "email_id": "0b1e3b0c-6a43-4d5e-9f3e-2f6c1f0a7d21",
            "status": "Delivered"
        }
    ]
}
```

## Known Issues

- MailCollector does not work with Outlook clients set with 'Try the new Outlook' set to on.
- Can error when accessing a given mail item actively in use in Outlook.
- MailCollector fails to read mailboxes configured to use IMAP.

## License

MailCollector is released under the [MIT License](LICENSE).
