# MailCollector

This .NET 4.8 application allows for the 'offline' collection of delivr.to campaign results from a Windows endpoint running Microsoft Outlook.

MailCollector leverages COM to communicate with the Outlook client and retrieve results. It can triage both links and attachments and supports filtering for results by campaign ID, sender, recipient, and mailbox folder.

## Installation

1. Clone repository
2. In the project folder run `dotnet restore`, then `dotnet build` (alternatively, open and build the project in Visual Studio).

## Usage

### Configuration

MailCollector uses a TOML file (`config.toml`) to tailor configuration without requiring the tool to be recompiled with hardcoded values. Unless specified otherwise, MailCollector will use the `config.toml` file in the same directory. This configuration file can contain the following values:

```toml
campaignId = "d373b79e-bd9a-4e3a-89e7-a14bdd3df9e3"
senderAddress = "no-reply@delivrto.me"
recipientAddress = "test-mailbox@delivr.to"
folderName = "delivr.to"
```

At a minimum, the configuration file needs to contain the `campaignID` value. Which can be retrieved from delivr.to either retrospectively, or by scheduling a future campaign (relevant for 'monitor' mode, see below).

The other options are optional and achieve the following:

- `senderAddress` - Defaults to `no-reply@delivrto.me`, but can be changed for custom senders.
- `recipientAddress` - Defaults to the primary Outlook mailbox, but useful if you have more than one mailbox configured in Outlook.
- `folderName` - Defaults to the inbox, but can be filtered to a specific folder, e.g. if rules have been configured.

### Execution

MailCollector operates in two similar but distinct modes, `capture` and `monitor`.

In `capture` mode, MailCollector can be run retrospectively, after a campaign has completed, to retrieve results. In `monitor` mode, arriving emails are actively monitored and analysed. Upon exitting MailCollector, results are then written to an `output.json` log file.

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
[+] JSON Log Written to: output.json
```

> NOTE: For attachment analysis, MailCollector saves payloads to a new directory in %TEMP%. Given AV/EDR may quarantine files that were otherwise permitted by mail gateways, it is recommended that appropriate allowlisting is in place to prevent result processing issues.

### Output

Upon results processing completion an `output.json` file is created in the same directory as the MailCollector executable. An example of its output can be seen below:

```json
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
}
```

This `output.json` can be uploaded directly in the delivr.to `Campaign Details` UI in order to synchronise results. 

Notably several fields in this JSON are provided for utility to end users, and aren't required for processing in the UI. A minimal results JSON is as follows:

```json
{
    "email_id": "6e8bb035-c50c-47c8-b56c-48332805b666",
    "link": "https://delivrto.me/links/a6882f3b-d165-469a-987f-4a5f7a94f2aa",
},
{
    "email_id": "a17e6ee1-0b06-4fbc-bf1b-ff4b323ef69d",
    "hash": "f3f96be1a14f905195dbc0c1c610ddff"
}
```
