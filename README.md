<div align="center">

## Send Outlook message from script


</div>

### Description

This allows you to fill out and send an Outlook message from script. Uncomment 2 lines to create a sub that can be included in a larger script. Great for admins who routinely send reports. 'T Runstein
 
### More Info
 
The script, as shown, is all hardcoded, but you could easily call "inputbox", or read inputs off a form to fill in the items. If you need help adapting this to your needs, let me know - I'm happy to help you customize it.

As with all .vbs scripts, you must have the scripting runtime. You also need Outlook installed.

You can choose to preview by using the .display, or send without preview by using .send


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[T Runstein](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/t-runstein.md)
**Level**          |Beginner
**User Rating**    |4.9 (78 globes from 16 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/t-runstein-send-outlook-message-from-script__4-6480/archive/master.zip)





### Source Code

```
'Uncomment the sub and
'end sub lines to use this in a program.
'Leaving these commented will allow you
'to run this as a standalone script
'sub SendAttach()
'Open mail, adress, attach report
dim objOutlk	'Outlook
dim objMail	'Email item
dim strMsg
const olMailItem = 0
'Create a new message
	set objOutlk = createobject("Outlook.Application")
	set objMail = objOutlk.createitem(olMailItem)
	objMail.To = "t_a_r@email.msn.com"
	objMail.cc = "" 'Enter an address here to include a carbon copy; bcc is for blind carbon copy's
'Set up Subject Line
	objMail.subject = "I saw your code on Planet Source Code on " & cstr(month(now)) & "/" & cstr(day(now)) & "/" & cstr(year(now))
'Add the body
	strMsg = "Your code rocks!" & vbcrlf
	strMsg = strMsg & "I voted and gave you an excellent rating!"
'To add an attachment, use:
	'objMail.attachments.add("C:\MyAttachmentFile.txt")
	objMail.body = strMsg
	objMail.display 'Use this to display before sending, otherwise call objMail.Send to send without reviewing
'Clean up
set objMail = nothing
set objOutlk = nothing
'end sub
```

