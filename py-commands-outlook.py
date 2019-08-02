import sys
import os.path
import datetime
import traceback
import win32timezone
import win32com.client as win32


def retrieve_project_parameters():
    
    parameters = sys.argv

    parameters_number = parameters.index("-traces") if "-traces" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        traces = parameters[parameters_number]
    else:
        traces = ""

    parameters_number = parameters.index("-command") if "-command" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        command = parameters[parameters_number]
    else:
        command = ""

    parameters_number = parameters.index("-account") if "-account" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        account = parameters[parameters_number]
    else:
        account = ""

    parameters_number = parameters.index("-email") if "-email" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        email = parameters[parameters_number]
    else:
        email = ""

    parameters_number = parameters.index("-folder") if "-folder" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        folder = parameters[parameters_number]
    else:
        folder = ""

    parameters_number = parameters.index("-by") if "-by" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        by = parameters[parameters_number]
    else:
        by = ""

    parameters_number = parameters.index("-to") if "-to" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        to = parameters[parameters_number]
    else:
        to = ""

    parameters_number = parameters.index("-cc") if "-cc" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        cc = parameters[parameters_number]
    else:
        cc = ""

    parameters_number = parameters.index("-bcc") if "-bcc" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        bcc = parameters[parameters_number]
    else:
        bcc = ""

    parameters_number = parameters.index("-subject") if "-subject" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        subject = parameters[parameters_number]
    else:
        subject = ""

    parameters_number = parameters.index("-body") if "-body" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        body = parameters[parameters_number]
    else:
        body = ""

    parameters_number = parameters.index("-read") if "-read" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        read = parameters[parameters_number]
    else:
        read = ""

    parameters_number = parameters.index("-attachments") if "-attachments" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        attachments = parameters[parameters_number]
    else:
        attachments = ""

    parameters_number = parameters.index("-path") if "-path" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        path = parameters[parameters_number]
    else:
        path = ""

    parameters_number = parameters.index("-draft") if "-draft" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        draft = parameters[parameters_number]
    else:
        draft = ""
        
    return {
        "traces": traces,
        "command": command,
        "account": account,
        "email": email,
        "folder": folder,
        "by": by,
        "to": to,
        "cc": cc,
        "bcc": bcc,
        "subject": subject,
        "body": body,
        "read": read,
        "attachments": attachments,
        "path": path,
        "draft": draft,
    }


def validate_project_parameters(parameters):
    
    traces = parameters["traces"]
    command = parameters["command"]
    account = parameters["account"]
    email = parameters["email"]
    folder = parameters["folder"]
    by = parameters["by"]
    to = parameters["to"]
    cc = parameters["cc"]
    bcc = parameters["bcc"]
    subject = parameters["subject"]
    body = parameters["body"]
    read = parameters["read"]
    attachments = parameters["attachments"]
    path = parameters["path"]
    draft = parameters["draft"]
    
    if traces == "" or traces.upper() == "FALSE":
        traces = False
    elif traces.upper() == "TRUE":
        traces = True
    else:
        return "ERROR: Invalid traces parameter! Parameter = " + str(traces)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Parameters retrieved start * ===")

    if command.upper() == "SEND":
        command = "SEND"
    elif command.upper() == "GET":
        command = "GET"
    elif command.upper() == "READ":
        command = "READ"
    elif command.upper() == "MOVE":
        command = "MOVE"
    elif command.upper() == "DELETE":
        command = "DELETE"
    elif command.upper() == "MARK":
        command = "MARK"
    elif command.upper() == "SAVE":
        command = "SAVE"
    elif command.upper() == "REPLY":
        command = "REPLY"
    elif command.upper() == "FORWARD":
        command = "FORWARD"
    elif command.upper() == "ATTACHMENTS":
        command = "ATTACHMENTS"
    elif not command == "":
        return "ERROR: Invalid command parameter! Parameter = " + str(command)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tCommand = " + str(command))

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAccount = " + str(account))
        
    if command.upper() == "MOVE":
        if email == "":
            return "ERROR: Invalid email parameter! Parameter = " + str(email)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tEmail = " + str(email))

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tFolder = " + str(folder))

    if by.upper() is not "" and not "@" in by:
        return "ERROR: Invalid by parameter! Parameter = " + str(by)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tCc = " + str(cc))

    if command.upper() == "SEND":
        if to.upper() == "" or not "@" in to:
            return "ERROR: Invalid to parameter! Parameter = " + str(to)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tTo = " + str(to))

    if cc.upper() is not "" and not "@" in cc:
        return "ERROR: Invalid cc parameter! Parameter = " + str(cc)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tCc = " + str(cc))

    if bcc.upper() is not "" and not "@" in bcc:
        return "ERROR: Invalid bcc parameter! Parameter = " + str(bcc)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tBcc = " + str(bcc))
        
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSubject = " + str(subject))
        
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tBody = " + str(body))
        
    if read is not "":
        if read.upper() == "TRUE":
            read = True
        elif read.upper() == "FALSE":
            read = False
        else:
            return "ERROR: Invalid read parameter! Parameter = " + str(read)
            
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tRead = " + str(read))

    if command.upper() == "SEND" and attachments is not "":
        attachments = attachments.replace(";", ",")
        attachments = attachments.split(",")
        attachments = [attachment.strip() for attachment in attachments]
        
        for attachment in attachments:
            if not os.path.isfile(attachment):
                return "ERROR: File in attachments does not exist! File = " + str(attachment)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttachments = " + str(attachments))
        
    if path is not "":
        if command.upper() == "SAVE" and not os.path.exists(os.path.dirname(path)):
            return "ERROR: The specified save path does not exist! Save path = " + str(os.path.dirname(path))
        elif command.upper() == "SAVE" and not str(os.path.basename(path)).endswith('.msg'):
            return "ERROR: The specified filename must end with '.msg'! Filename = " + str(os.path.basename(path))
        
        if command.upper() == "ATTACHMENTS" and not os.path.exists(path):
            return "ERROR: The specified save path does not exist! Save path = " + str(path)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tDownload path = " + str(path))
    
    if draft is not "":
        if draft.upper() == "TRUE":
            draft = True
        elif draft.upper() == "FALSE":
            draft = False
        else:
            return "ERROR: Invalid draft parameter! Parameter = " + str(draft)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tDraft = " + str(draft))

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Parameters retrieved end * ===")
        
    return {
        "traces": traces,
        "command": command,
        "account": account,
        "email": email,
        "folder": folder,
        "by": by,
        "to": to,
        "cc": cc,
        "bcc": bcc,
        "subject": subject,
        "body": body,
        "read": read,
        "attachments": attachments,
        "path": path,
        "draft": draft,
    }
    
    
def get_emails(emails):
    
    data = []
    data.append(['EntryID', 'ReceivedDate', 'ReceivedTime', 'UnRead', 'Sender', 'Cc', 'Subject', 'AttachmentsCount', 'Attachments'])
    
    for email in emails:
        email_info = []
        
        email_info.append(email.EntryID)
        email_info.append(str(email.ReceivedTime.month) + '/' + str(email.ReceivedTime.day) + '/' + str(email.ReceivedTime.year))
        email_info.append(str(email.ReceivedTime.hour).zfill(2) + ':' + str(email.ReceivedTime.minute).zfill(2) + ':' + str(email.ReceivedTime.second).zfill(2))
        email_info.append(not email.UnRead)
        email_info.append(str(email.Sender) + " (" + str(email.SenderEmailAddress) + ")")
        email_info.append(email.Cc)
        email_info.append(email.Subject)
        email_info.append(email.Attachments.Count)
        attachments = []
        for attachment in email.Attachments:
            attachments.append(attachment.DisplayName)
        email_info.append(attachments)
        
        data.append(email_info)
    
    return data


def send_email(outlook_msg, to, cc, bcc, subject, body, attachments, draft, traces, reply=False, forward=False):
    if not reply and not forward:
        outlook_msg.Subject = subject
    
    if not reply:
        outlook_msg.To = to
        outlook_msg.CC = cc
        outlook_msg.BCC = bcc        

    outlook_msg.GetInspector()
    
    outlook_msg_signature = outlook_msg.HTMLBody
    outlook_msg.HTMLBody = "<BODY>" + body + "</BODY>" + outlook_msg_signature

    for attachment in attachments:
        outlook_msg.Attachments.Add(attachment)

    if not draft:
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to send email...")
            
        outlook_msg.Send()
        
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tEmail sent!")
            
    else:
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to save email as draft...")
            
        outlook_msg.Save()
        
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tEmail saved as draft!")
            
    return outlook_msg


def execute_command(parameters):
    
    traces = parameters["traces"]
    command = parameters["command"]
    account = parameters["account"]
    email = parameters["email"]
    folder = parameters["folder"]
    by = parameters["by"]
    to = parameters["to"]
    cc = parameters["cc"]
    bcc = parameters["bcc"]
    subject = parameters["subject"]
    body = parameters["body"]
    read = parameters["read"]
    attachments = parameters["attachments"]
    path = parameters["path"]
    draft = parameters["draft"]
    
    outlook = win32.Dispatch('Outlook.Application')
    outlook_accounts = outlook.Session.Accounts
    outlook_accounts_list = [outlook_account.DisplayName for outlook_account in outlook_accounts]
    
    if account is not "":
        if not account in outlook_accounts_list:
            return f"ERROR: The specified account ({account}) is not one of the available accounts ({outlook_accounts_list})"
    
    try:

        if command.upper() == "SEND":
            if draft is "":
                draft = False
            
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Sending email start (draft: {str(draft)}) * ===")
                    
            outlook_msg = outlook.CreateItem(0)
            
            if account is not "":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to select account: " + account)
                    
                for outlook_account in outlook_accounts:
                    if account == outlook_account.DisplayName:
                        outlook_msg._oleobj_.Invoke(*(64209, 0, 8, 0, outlook_account))     # https://stackoverflow.com/a/35908213
                        
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSelecting account complete!")
                            
                        break
                        
            outlook_msg = send_email(outlook_msg, to, cc, bcc, subject, body, attachments, draft, traces)
                
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Sending email end (draft: {str(draft)}) * ===")
            
        else:
            if account is not "":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to select account: " + account)
                    
                outlook_namespace = None
                    
                for outlook_account in outlook_accounts:
                    if account == outlook_account.DisplayName:
                        outlook_namespace = outlook_account.DeliveryStore
                        
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSelecting account complete!")
                            
                        break
                    
                if outlook_namespace is None:
                    return "ERROR: Failed to select the specified Outlook account!"
                    
            else:                    
                outlook_namespace = outlook.GetNamespace('MAPI')
                
            outlook_folder = outlook_namespace.GetDefaultFolder(6)
            
            if command.upper() == "GET":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Getting emails start * ===")
                
                if folder is not "":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to select folder: " + folder)
                        
                    folders = folder.split("\\")
                    for folder in folders:
                        outlook_folder = outlook_folder.Folders(folder)
                        
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSelecting folder complete!")
                    
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tLooping through emails in the folder...")
                        
                emails = get_emails(outlook_folder.Items)
                print(emails)
                    
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAll emails retrieved!")
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Getting emails end * ===")
                
            else: 
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to retrieve email with id: " + email)
                    
                outlook_email = outlook_namespace.GetItemFromID(email)
                
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tEmail successfully retrieved: " + str(outlook_email))
                
                if command.upper() == "READ":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to read email body...")
                        
                    print(outlook_email.Body)
                    
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tEmail body read!")
                
                elif command.upper() == "MOVE":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Moving email start * ===")
                
                    if folder is not "":
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to select folder: " + folder)
                            
                        folders = folder.split("\\")
                        for folder in folders:
                            outlook_folder = outlook_folder.Folders(folder)
                            
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSelecting folder complete!")
                            
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to move the email...")
                    
                    outlook_email.Move(outlook_folder)
                            
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tMoving email complete!")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Moving email end * ===")
                    
                elif command.upper() == "DELETE":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Deleting email start * ===")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to delete the email...")
                        
                    outlook_email.Delete()
                            
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tDeleting email complete!")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Deleting email end * ===")
                    
                elif command.upper() == "MARK":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Marking email start * ===")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tAttempting to mark the email {'read' if read else 'unread'}...")
                        
                    outlook_email.UnRead = not read
                            
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tDeleting email complete!")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Deleting email end * ===")
                    
                elif command.upper() == "SAVE":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Saving email start * ===")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tAttempting to save the email as {str(path)}")
                        
                    outlook_email.SaveAs(path)
                            
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSaving email complete!")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"=== * Saving email end * ===")
                    
                elif command.upper() == "REPLY":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Replying email start * ===")
                    
                    outlook_msg = outlook_email.ReplyAll()
                    outlook_msg = send_email(outlook_msg, to, cc, bcc, subject, body, attachments, draft, traces, reply=True)
                    
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Replying email end * ===")
                    
                elif command.upper() == "FORWARD":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Forward email start * ===")
                    
                    outlook_msg = outlook_email.Forward()
                    outlook_msg = send_email(outlook_msg, to, cc, bcc, subject, body, attachments, draft, traces, forward=True)
                    
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Forward email end * ===")
                    
                elif command.upper() == "ATTACHMENTS":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Saving attachments start * ===")
                        
                    for outlook_email_attachment in outlook_email.Attachments:
                        save_path = os.path.join(path, outlook_email_attachment.FileName)
                            
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tAttempting to save the attachment: {save_path}")
                        
                        outlook_email_attachment.SaveAsFile(save_path)
                            
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tAttachment saved!")
                        
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Saving attachments end * ===")
                
    except:
        outlook_msg = None
        outlook_namespace = None
        outlook_folder = None
        outlook_email = None
        
        outlook_accounts = None
        outlook = None
        
        print(traceback.format_exc())
        return "ERROR: Unexpected issue!"
    
    outlook_msg = None
    outlook_namespace = None
    outlook_folder = None
    outlook_email = None
    
    outlook_accounts = None
    outlook = None
    
    return True
    
    
def main():
    
    parameters = retrieve_project_parameters()
    
    parameters = validate_project_parameters(parameters)
    if not isinstance(parameters, dict):
        print(str(parameters))
        return
    
    valid = execute_command(parameters)
    if not valid is True:
        print(str(valid))
        return
    
    print("SUCCESS")
    
    
if __name__ == "__main__":
    main()
