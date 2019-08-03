# py-commands-outlook

<i>Still under development...</i>

This program allows you to execute Outlook commands such as sending emails, getting all emails in a folder, reading emails, moving emails, and more. This is especially useful if you do not have access or is allowed to perform email actions using SMTP and IMAP server connections. You can run the program via the CMD or as part of an automation script in an RPA tool like Foxtrot. This solution is meant to supplement Foxtrots core email functionality and enable you to perform email automation via your local Outlook application instead of email server connections. The solution is written in Python using the module "pywin32". You can see the [full source code here](https://github.com/foxtrot-alliance/py-commands-outlook/blob/master/py-commands-outlook.py).

## Outlook warning

When using this program, you might experience a warning from Outlook or potentially be fully prevented from using the program. If so, please read [this official support article by Microsoft](https://support.microsoft.com/en-us/help/3189806/a-program-is-trying-to-send-an-e-mail-message-on-your-behalf-warning-i).

## Installation

1. Download the [latest version](https://github.com/foxtrot-alliance/py-commands-outlook/releases/download/v0.0.4/py-commands-outlook_v0.0.4.zip).
2. Unzip the folder somewhere appropriate, we suggest directly on the C: drive for easier access. So, your path would be similar to "C:\py-commands-outlook_v0.0.4".
3. After unzipping the files, you are now ready to use the program. The only file you will have to be concerned about is the actual .exe file in the folder, however, all the other files are required for the solution to run properly.
4. Open Foxtrot (or any other RPA tool) to set up your action. In Foxtrot, you can utilize the functionality of the program via the DOS Command action (alternatively, the Powershell action).

## Usage

When using the program via Foxtrot, the CMD, or any other RPA tool, you need to reference the path to the program exe file. If you placed the program directly on your C: drive as recommended, the path to your program will be similar to: 
```
C:\py-commands-outlook_v0.0.4\py-commands-outlook_v0.0.4.exe
```
TIP: Make sure NOT to surround the path with quotation marks in your commands.

## Commands

All the available commands are specified [here](#all-available-parameters). Note, all parameters surrounded by [-x "X"] means that they are optional. For a more detailed description of each command, read the [detailed command description section](#detailed-command-description).

The solution offers three main commands:
* Send email
```
PROGRAM_EXE_PATH -command "send" [-account "X"] -to "X" [-cc "X"] [-bcc "X"] -subject "X" -body "X" [-attachments "X"] [-draft "X"]
```
* Get all emails in a folder
```
PROGRAM_EXE_PATH -command "get" [-account "X"] -path "X" -delimiter "X"
```
* Work with the email based on the ID retrieved from getting all the emails
```
PROGRAM_EXE_PATH -command "X" -email "X" [-to "X"] [-cc "X"] [-bcc "X"] [-subject "X"] [-body "X"] [-attachments "X"] [-draft "X"] [-folder "X"] [-path "X"] [-unread "X"]
```

The solution will give an output to the selected variable in the DOS Command action to indicate whether the command was executed successfully or not.

## Examples

### Sending emails

These two examples will send emails. The first one will send a simple email using the default Outlook account, the second will select a different account to send from and include attachments. Notice that you can use HTML in the email body.
```
PROGRAM_EXE_PATH -command "send" -to "mbalslow@basico.dk" -subject "Sending emails!" -body "Pretty cool, right?!"
PROGRAM_EXE_PATH -command "send" -account "mbalslow@foxtrotalliance.com" -to "mbalslow@basico.dk; mbalslow@gmail.com" -subject "Sending emails to multiple people!" -body "Also pretty cool, right?!" -attachments "C:\file1.txt, C:\file2.txt"
```

### Getting emails

These two examples will get all emails. By default, this command will look in the main email inbox. If you wish to look in a subfolder of the inbox, you can specify this. The output will be stored in a CSV file using the specified delimiter. Again, if you have multiple accounts, you can specify which one you wish to use.
```
PROGRAM_EXE_PATH -command "get" -path "c:\emails.csv" -delimiter ","
PROGRAM_EXE_PATH -command "get" -account "mbalslow@foxtrotalliance.com" -path "c:\emails.csv" -delimiter ";"
```

### Working with emails

After getting the emails, you can work with them referencing the email id. IMPORTANT: Keep in mind that the email id changes when the email is moved to a different folder! Again, if you have multiple accounts, you can specify which one you wish to use.
```
PROGRAM_EXE_PATH -command "delete" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000"
PROGRAM_EXE_PATH -command "move" -account "mbalslow@foxtrotalliance.com" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -folder "Notes\GitHub"
```

## All available parameters
```
-command: "send", "get", "read", "move", "delete", "mark", "save", "reply", "forward", "attachments"
  This is the command you wish the solution to execute.

-account: "X", default = none (will use the default Outlook account if not specified)
  This is the account to work with in Outlook.

-to: "X", required if "send" or "forward" command is used.
  This is the email address(es) of the receiver(s) of the email to be sent.

-cc: "X", default = none, relevant if "send" or "forward" command is used.
  This is the email address(es) of the cc(s) of the email to be sent.

-bcc: "X", default = none, relevant if "send" or "forward" command is used.
  This is the email address(es) of the bcc(s) of the email to be sent.

-subject: "X", required if "send" command is used.
  This is the email subject of the email to be sent.

-body: "X", required if "send", "reply", or "forward" command is used.
  This is the email body of the email to be sent. NOTE: You can use HTML to format your email body.

-attachments: "X", default = none, relevant if "send", "reply", or "forward" command is used.
  This is the file paths of the attachments to be added to the email to be sent.

-draft: "true"/"false", default = "false", relevant if "send", "reply", or "forward" command is used.
  This determines whether you wish to save your email as a draft instead of sending it.

-folder: "X", default = none, relevant if "get" or "move" command is used.
  This is the subfolder path for the folder to either get all emails or move a selected email to.

-path: "X", default = none, required if "get", "save", or "attachments" command is used.
  This is the path used to define where the output files should be saved.

-delimiter: ","/";"/"|", required if "get" command is used.
  This is the parameter used to define the delimiter of the CSV output of the "get" command.

-unread: "true"/"false", required if "mark" command is used.
  This is the parameter used to define whether the specified email should be marked unread or read.

-traces: "true"/"false", default = "false"
  This determines whether you wish the output to include traces, information about the execution.
```

## Detailed command description

### Send email
Parameters:
```
PROGRAM_EXE_PATH -command "send" [-account "X"] -to "X" [-cc "X"] [-bcc "X"] -subject "X" -body "X" [-attachments "X"] [-draft "X"]
```
Examples:
```
PROGRAM_EXE_PATH -command "send" -to "mbalslow@basico.dk" -subject "Sending emails!" -body "Pretty cool, right?!"
PROGRAM_EXE_PATH -command "send" -account "mbalslow@foxtrotalliance.com" -to "mbalslow@basico.dk; mbalslow@gmail.com" -subject "Sending emails to multiple people!" -body "Also pretty cool, right?!" -attachments "C:\file1.txt, C:\file2.txt"
```

### Get emails
Parameters:
```
PROGRAM_EXE_PATH -command "get" [-account "X"] -path "X" -delimiter "X"
```
Examples:
```
PROGRAM_EXE_PATH -command "get" -path "c:\emails.csv" -delimiter ","
PROGRAM_EXE_PATH -command "get" -account "mbalslow@foxtrotalliance.com" -path "c:\emails.csv" -delimiter ";"
```

### Read email
Parameters:
```
PROGRAM_EXE_PATH -command "read" [-account "X"] -email "X" [-path "X"]
```
Examples:
```
PROGRAM_EXE_PATH -command "read" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000"
PROGRAM_EXE_PATH -command "read" -account "mbalslow@foxtrotalliance.com" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -path "c:\email.txt"
```

### Move email
Parameters:
```
PROGRAM_EXE_PATH -command "move" [-account "X"] -email "X" -folder "X"
```
Examples:
```
PROGRAM_EXE_PATH -command "move" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -folder "Notes"
PROGRAM_EXE_PATH -command "move" -account "mbalslow@foxtrotalliance.com" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -folder "Notes\GitHub"
```

### Delete email
Parameters:
```
PROGRAM_EXE_PATH -command "delete" [-account "X"] -email "X"
```
Examples:
```
PROGRAM_EXE_PATH -command "delete" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000"
PROGRAM_EXE_PATH -command "delete" -account "mbalslow@foxtrotalliance.com" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000"
```

### Mark email
Parameters:
```
PROGRAM_EXE_PATH -command "mark" [-account "X"] -email "X" -unread "X"
```
Examples:
```
PROGRAM_EXE_PATH -command "mark" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -mark "false"
PROGRAM_EXE_PATH -command "mark" -account "mbalslow@foxtrotalliance.com" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -mark "true"
```

### Save email
Parameters:
```
PROGRAM_EXE_PATH -command "save" [-account "X"] -email "X" -path "X"
```
Examples:
```
PROGRAM_EXE_PATH -command "save" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -path "c:\email.msg"
PROGRAM_EXE_PATH -command "save" -account "mbalslow@foxtrotalliance.com" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -path "c:\email.msg"
```

### Reply email
Parameters:
```
PROGRAM_EXE_PATH -command "reply" [-account "X"] -email "X" -body "X" [-attachments "X"] [-draft "X"]
```
Examples:
```
PROGRAM_EXE_PATH -command "reply" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -body "That sounds great, thank you for the email!"
PROGRAM_EXE_PATH -command "reply" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -body "Cool, please see my attached file" -attachments "c:\file.txt" -draft "true"
```

### Forward email
Parameters:
```
PROGRAM_EXE_PATH -command "forward" [-account "X"] -email "X" -to "X" [-cc "X"] [-bcc "X"] -body "X" [-attachments "X"] [-draft "X"]
```
Examples:
```
PROGRAM_EXE_PATH -command "forward" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -to "mbalslow@basico.dk" -body "FYI"
PROGRAM_EXE_PATH -command "forward" -account "mbalslow@foxtrotalliance.com" -email "00000000F7888E77BE28594189F5D1712F48C6CB0700127A6B9C0A128644AB22BF3FAB4C3A59000000F56BDA0000592EAC7355CF234FB22B18BB604DD20900007AC2858C0000" -to "mbalslow@basico.dk" -cc "mbalslow@gmail.com" -body "I have added another file" -attachments "c:\file.txt" -draft "true"
```
