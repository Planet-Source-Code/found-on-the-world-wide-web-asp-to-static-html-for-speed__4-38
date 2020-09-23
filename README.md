<div align="center">

## ASP to static HTML for speed


</div>

### Description

If you have a large amount of data to give to the user as HTML and this data needs to change once a day then this will speed up the process for the user.

The following code will create a file the first time a page is hit for each day.

The upside of doing it this way is you have a record of what the use saw on any given day.

The downside is the first person takes the performance hit to write the page and you need to check to make sure the user came to this page first. In other words, if they save yesterdays page as a fovorite then they will see old data unless you redirect.

I used the month and day to handle this problem. I did not use the year. There are many other ways to handle this problem.

http://www.truegeeks.com/asp/mam/osdoc/osframe.asp
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Intermediate
**User Rating**    |3.0 (12 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-asp-to-static-html-for-speed__4-38/archive/master.zip)





### Source Code

```
Dim fs, fsmyfile, todayfile, ckdayfile, cr, qt
'Get name of file as it needs to be today
todayfile="Cur"&cstr(month(date()))&cstr(day(date()))
ckdayfile=""&cstr(month(date()))&cstr(day(date()))&""
todayfile=trim(todayfile)&".asp"
'Create FileSystemObject
Set fs = CreateObject("Scripting.FileSystemObject")
'File may not be built
On Error Resume Next
'Check to see if we already have the HTML file Built
Set fsmyfile = fs.OpenTextFile("c:\inetpub\scripts\asp\jeff\"+todayfile,1,0)
if err<>0 then		'Need to build today
	fsmyfile.Close	'Close File
	Set fsmyfile = fs.OpenTextFile("c:\inetpub\scripts\asp\jeff\"+todayfile,8,1,0)
	cr=chr(13)	'Save some typing (I'm lazy)
	qt=chr(34)	'The Only way I could get the quote marks correct
	codeout="<%@ LANGUAGE=""VBSCRIPT"" %"&">"&cr
	codeout=codeout&"<%"&cr
	codeout=codeout&"today="&qt&cstr(month(date()))&cstr(day(date()))&qt&cr
	codeout=codeout&cr&"if today<>"&qt&ckdayfile&qt&" then"&cr
	codeout=codeout&"response.redirect("&qt&"wrtest.asp"&qt&")"&cr
	codeout=codeout&"else %"&">"&cr
	fsmyfile.Writeline(""&codeout&cr&_
	"<HTML>"&cr&_
	"<title>Write and Check Raw HTML For Speed</title>"&cr&_
	"<BODY>"&cr&_
	"Hello todays file is called "&todayfile&cr&_
	"</BODY>"&cr&_
	"</HTML>"&cr&_
	"<"&"%End if"&cr&_
	"%"&">")
	fsmyfile.close
	fs.close
	Response.Redirect(todayfile)	'Send them to new file
else
	fsmyfile.close
	fs.Close
	Response.Redirect(todayfile)	'Send them to current file
end if%>
```

