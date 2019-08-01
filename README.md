# py-commands-outlook

## Installation
[Download](https://github.com/foxtrot-alliance/py-commands-outlook/releases/download/v0.0.1/py-commands-outlook_v0.0.1.zip)

## Parameters
* [-traces "true"/"false"]
* -command "send"
* -account "x@y.com"
* [-folder "xyz"] (N/A)
* [-by "xyz"] (N/A)
* -to "xyz; xyz; xyz"
* [-cc "xyz; xyz; xyz"]
* [-bcc "xyz; xyz; xyz"]
* -subject "xyz"
* -body "xyz"
* -attachments "c:\test.txt, c:\test2.txt"
* [-download_path "c:\"] (N/A)
* [-draft" "true"/"false"]

## Example
```
EXE_PATH -command "send" -to "mbalslow@foxtrotalliance.com" -cc "mbalslow@gmail.com" -subject "Sending emails!" -body "Pretty cool,<br>Right?!" -attachments "C:\test1.png, C:\test2.png" -account "mbalslow@foxtrotalliance.com"
```
