import sys
import os.download_path
import datetime
import traceback
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

    parameters_number = parameters.index("-attachments") if "-attachments" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        attachments = parameters[parameters_number]
    else:
        attachments = ""

    parameters_number = parameters.index("-download_path") if "-download_path" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        download_path = parameters[parameters_number]
    else:
        download_path = ""

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
        "folder": folder,
        "by": by,
        "to": to,
        "cc": cc,
        "bcc": bcc,
        "subject": subject,
        "body": body,
        "attachments": attachments,
        "download_path": download_path,
        "draft": draft,
    }


def validate_project_parameters(parameters):
    
    traces = parameters["traces"]
    command = parameters["command"]
    account = parameters["account"]
    folder = parameters["folder"]
    by = parameters["by"]
    to = parameters["to"]
    cc = parameters["cc"]
    bcc = parameters["bcc"]
    subject = parameters["subject"]
    body = parameters["body"]
    attachments = parameters["attachments"]
    download_path = parameters["download_path"]
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
    elif command.upper() == "GET_ALL":
        command = "GET_ALL"
    elif not command == "":
        return "ERROR: Invalid command parameter! Parameter = " + str(command)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tCommand = " + str(command))

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAccount = " + str(account))

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

    if command.upper() == "SEND" and attachments is not "":
        attachments = attachments.replace(";", ",")
        attachments = attachments.split(",")
        attachments = [attachment.strip() for attachment in attachments]
        
        for attachment in attachments:
            if not os.path.isfile(attachment):
                return "ERROR: File in attachments does not exist! File = " + str(attachment)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttachments = " + str(attachments))
        
    if download_path is not "" and not os.path.exists(download_path):
        return "ERROR: The specified download path does not exist! Download path = " + str(download_path)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tDownload path = " + str(download_path))
    
    if draft == "" or draft.upper() == "FALSE":
        draft = False
    elif draft.upper() == "TRUE":
        draft = True
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
        "folder": folder,
        "by": by,
        "to": to,
        "cc": cc,
        "bcc": bcc,
        "subject": subject,
        "body": body,
        "attachments": attachments,
        "download_path": download_path,
        "draft": draft,
    }

    
def execute_command(parameters):
    
    traces = parameters["traces"]
    command = parameters["command"]
    account = parameters["account"]
    folder = parameters["folder"]
    by = parameters["by"]
    to = parameters["to"]
    cc = parameters["cc"]
    bcc = parameters["bcc"]
    subject = parameters["subject"]
    body = parameters["body"]
    attachments = parameters["attachments"]
    download_path = parameters["download_path"]
    draft = parameters["draft"]
    
    try:

        if "SEND" in command.upper():
            
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Sending email start * ===")
            
            outlook = win32.Dispatch('Outlook.Application')
            outlook_accounts = outlook.Session.Accounts
            outlook_accounts_list = [outlook_account.DisplayName for outlook_account in outlook_accounts]
            
            if account is not "":
                if not account in outlook_accounts_list:
                    return f"ERROR: The specified account ({account}) is not one of the available accounts ({outlook_accounts_list})"
                    
            msg = outlook.CreateItem(0)
            
            if account is not "":
                for outlook_account in outlook_accounts:
                    if account == outlook_account.DisplayName:
                        msg._oleobj_.Invoke(*(64209, 0, 8, 0, outlook_account))     # https://stackoverflow.com/a/35908213
                        
            msg.To = to
            msg.CC = cc
            msg.BCC = bcc
            msg.Subject = subject
            
            msg.GetInspector()
            signature = msg.HTMLBody
            msg.HTMLBody = "<BODY>" + body + "</BODY>" + signature
            
            for attachment in attachments:
                msg.Attachments.Add(attachment)

            if not draft:
                msg.Send()
            else:
                msg.Save()
            
            msg = None
            outlook = None
            outlook_accounts = None
            
        else:
            
            pass
    
    except:
        print(traceback.format_exc())
        return "ERROR: Unexpected issue!"
    
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
