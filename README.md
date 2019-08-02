# py-commands-outlook

<i>Still under development...</i>

## Installation
[Download](https://github.com/foxtrot-alliance/py-commands-outlook/releases/download/v0.0.3/py-commands-outlook_v0.0.3.zip)

## Parameters
* -traces "true"/"false"
* -command "send","get", "read", "move", "delete", "mark", "save", "reply", "forward", "attachments"
* -account "x@y.com"
* -folder "xyz"
* -by "xyz"
* -to "xyz; xyz; xyz"
* -cc "xyz; xyz; xyz"
* -bcc "xyz; xyz; xyz"
* -subject "xyz"
* -body "xyz"
* -attachments "c:\test.txt, c:\test2.txt"
* -read "true"/"false"
* -path "c:\"
* -draft" "true"/"false"

## Example
```
EXE_PATH -command "send" -to "mbalslow@foxtrotalliance.com" -cc "mbalslow@gmail.com" -subject "Sending emails!" -body "Pretty cool,<br>Right?!" -attachments "C:\test1.png, C:\test2.png"
```
