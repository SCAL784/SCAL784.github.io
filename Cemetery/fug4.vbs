Set FSO=CreateObject("Scripting.FileSystemObject")

Set ie = CreateObject("InternetExplorer.Application")
'ie.Visible = True
ie.Navigate "file:///D:/Dropbox/cma/Cemetery/Fugate%20Cemetery.html"
While ie.Busy : WScript.Sleep 100 : Wend

' How to write file
outFile="names.csv"
Set objFile = FSO.CreateTextFile(outFile,True)
'objFile.Write "test,string" & vbCrLf



'FSO.OpenTextFile("debug.html", 2, True).Write ie.document.body.innerHtml

For Each tr In ie.document.getElementsByTagName("tr")
'  If InStr(tr.innerText, "RegistrationDTO.register") > 0 Then
    Set row = tr
'  End If
objFile.Write row.children(0).innerText & "," & row.children(1).innerText & "," & row.children(2).innerText & "," & row.children(3).innerText & vbCrLf
'WScript.Echo row.children(1).innerText
Next

objFile.Close
ie.Quit
WScript.Echo "Done"
Wscript.Quit(0)

'Set daTable = Browser.Document.getElementById("Table")
'strOut = ""
'For i = 0 To daTable.rows.length - 1
'  For j = 0 To daTable.rows.item(i).cells.length - 1
'    strOut = strOut & daTable.rows.item(i).cells.item(j).innerText & Chr(9)
'  Next
'  strOut = strOut & Chr(13) & Chr(10)
'Next
'Browser.Quit
'Browser = Null
'Msgbox "Table Data: " & Chr(13) & Chr(10) & strOut
