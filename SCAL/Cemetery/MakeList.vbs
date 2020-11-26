Set FSO=CreateObject("Scripting.FileSystemObject")

outFile="names.csv"
Set FileO = FSO.CreateTextFile(outFile,True)

inFile = "c.txt"
Set FileI = FSO.OpenTextFile(inFile)

Do Until FileI.AtEndOfStream
  strLine = FileI.ReadLine
  fil = split(strLine, ".")
  fname = "file:///D:/Dropbox/cma/Cemetery/" & strLine
  Set ie = CreateObject("InternetExplorer.Application")
  ie.Navigate fname

  While ie.Busy : WScript.Sleep 100 : Wend

  For Each tr In ie.document.getElementsByTagName("tr")
    Set row = tr
    FileO.Write fil(0) & "," & row.children(0).innerText & "," & row.children(1).innerText & "," & row.children(2).innerText & "," & row.children(3).innerText & vbCrLf
  Next

  ie.Quit
  WScript.Sleep 100
Loop

FileI.Close
FileO.Close

WScript.Echo "Done"
Wscript.Quit(0)

