<%
' Output XML
Response.ContentType = "text/xml"

' Functions borrowed from http://www.webicy.com/scripts-content-management/8076-asp-tutorial-converting-normal-date-into-rfc-822-format-date-using-asp.html
Function return_RFC822_Date(myDate, offset)
   Dim myDay, myDays, myMonth, myYear
   Dim myHours, myMonths, mySeconds

   myDate = CDate(myDate)
   myDay = WeekdayName(Weekday(myDate),true)
   myDays = Day(myDate)
   myMonth = MonthName(Month(myDate), true)
   myYear = Year(myDate)
   myHours = zeroPad(Hour(myDate), 2)
   myMinutes = zeroPad(Minute(myDate), 2)
   mySeconds = zeroPad(Second(myDate), 2)

   return_RFC822_Date = myDay&", "& _
                                  myDays&" "& _
                                  myMonth&" "& _ 
                                  myYear&" "& _
                                  myHours&":"& _
                                  myMinutes&":"& _
                                  mySeconds&" "& _ 
                                  offset
End Function 
Function zeroPad(m, t)
   zeroPad = String(t-Len(m),"0")&m
End Function

' Write the XML data
Response.Write("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>")
Response.Write("<rss version=""2.0"">")
Response.Write("<channel>")
Response.Write("<title>New Plymouth Reformed Church Sermons</title>")
Response.Write("<link>https://sermons.azurewebsites.net/</link>")

' Start file system access
Set fs = CreateObject("Scripting.FileSystemObject")

' Set the name of the current folder
Set folder = fs.GetFolder(Request.ServerVariables("APPL_PHYSICAL_PATH"))

' Get all of the files in the current folder
Set files = folder.Files

' Display each file one by one
For Each file In files
	' If we have an MP3 or PDF file, then display an item for it
	If LCase(Right(file.Name, 4)) = ".mp3" Or LCase(Right(file.Name, 4)) = ".pdf" Then
		' The RSS item
		Response.Write("<item>")
		Response.Write("<title>")
		Response.Write(Left(file.Name, Len(file.Name) - 4))
		Response.Write("</title>")
		Response.Write("<pubDate>")
		Response.Write(return_RFC822_Date(file.DateLastModified, "+1200"))
		Response.Write("</pubDate>")
		Response.Write("<link>")
		Response.Write("https://sermons.azurewebsites.net/")
		Response.Write(file.Name)
		Response.Write("</link>")
		Response.Write("</item>")
	End If
Next

' End the data
Response.Write("</channel></rss>")

' Clean up
Set files = Nothing
Set folder = Nothing
Set fs = Nothing
%>