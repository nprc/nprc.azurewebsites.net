<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Plymouth Reformed Church Sermon Archive</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link type="image/vnd.microsoft.icon" rel="shortcut icon" href="https://sermons.azurewebsites.net/favicon.ico">
<link rel="alternate" href="https://sermons.azurewebsites.net/RSS.asp" type="application/rss+xml" title="New Plymouth Reformed Church Sermons" id="rss">
<meta name="description" content="New Plymouth Reformed Church MP3 Sermons">
<meta name="keywords" content="MP3 Sermons New Plymouth Christian Protestant Reformed Church">
<meta name="author" content="Peter Chapman">
</head>
<body>
<script type="text/javascript">
(function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
(i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
})(window,document,'script','//www.google-analytics.com/analytics.js','ga');
ga('create', 'UA-61493830-2', 'auto');
ga('send', 'pageview');
</script>
<h1>New Plymouth Reformed Church Sermon Archive</h1>
<p>
	<a href="https://sermons2015.azurewebsites.net/">2015</a> <a href="https://sermons2016.azurewebsites.net/">2016</a> <a href="https://sermons.azurewebsites.net/">2017</a>
</p>
<!--Do Not Modify The Following Code, Only Application("ICONS_PATH"), and Application("ICONS_VIRTUAL_PATH") If Needed-->
<%
' Set the path to where the file icons are kept if it is empty
' Note: You should set these in global.asa
' All icons in this folder must be in gif format
' For example, the icon for Word documents would be doc.gif
If Application("ICONS_VIRTUAL_PATH") = "" Then Application("ICONS_VIRTUAL_PATH") = "/icons" ' Manually set the virtual path here if you don't have a global variable
If Application("ICONS_PATH") = "" Then Application("ICONS_PATH") = "D:\home\site\wwwroot\icons" ' Manually set the path here if you don't have a global variable

' Start file system access
Set fs = CreateObject("Scripting.FileSystemObject")

' Set the name of the current folder
Set folder = fs.GetFolder(Request.ServerVariables("APPL_PHYSICAL_PATH"))

' Get all of the files in the current folder
Set files = folder.Files

' Display each file one by one
For Each filename In files
	' Get the file's extension
	Dim fileExtension, filenameArray
	filenameArray = Split(filename.Name, ".", -1, 1)
	fileExtension = LCase(filenameArray(UBound(filenameArray)))
	' If we have an MP3 or PDF file, then display a link to it
	If fileExtension = "mp3" Or fileExtension = "pdf" Then
		' Display the icon if one exists. This If statement speeds processing.
		If Application("ICONS_PATH") <> "" Then
			' If we have an image for this file type show it, otherwise show the default icon if it exists
			If fs.FileExists(Application("ICONS_PATH") + "\" + fileExtension + ".gif") Then
				Response.Write("<a href=""" + filename.Name + """><img src=""" + Application("ICONS_VIRTUAL_PATH") + "/" + fileExtension + ".gif"" align=""absmiddle"" border=""0"" alt=""" + Left(filename.Name, Len(filename.name) - Len(fileExtension) - 1) + """></a>")
			End If
		End If
		' Display the name of the file
		Response.Write(" <a href=""" + filename.Name + """>" + filename.Name + "</a><br>")
	End If
Next

' Clean up
Set files = Nothing
Set folder = Nothing
Set fs = Nothing
%>
<!--You Can Modify All Of The Code Below-->
<p align="right">&copy; 2015-2017 <a href="http://nprc.nz/">New Plymouth Reformed Church</a></p>
</body>
</html>

