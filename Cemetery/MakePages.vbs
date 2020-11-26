Const ForReading = 1 
Dim FileLines()
Redim FileLines(1000,5)
i = 0

Set FSO=CreateObject("Scripting.FileSystemObject")

InFile="names.csv"
Set FileIn = FSO.OpenTextFile(InFile,ForReading)

Do Until FileIn.AtEndOfStream
  InputStr = FileIn.ReadLine
  InputArr = Split(InputStr, ",")
  FileLines(i,0) = InputArr(0)
  FileLines(i,1) = InputArr(1)
  FileLines(i,2) = InputArr(2)
  FileLines(i,3) = InputArr(3)
  FileLines(i,4) = InputArr(4)

  If Len(InputArr(3)) > 1 Then
    If InStr(InputStr, "Name ,Era ") = 0 Then
      i = i + 1
    End If
  End If
Loop

FileIn.Close

for a = i - 1 To 0 Step -1
    for j = 0 to a
        if FileLines(j,0)>FileLines(j+1,0) then
            temp0=FileLines(j+1,0)
            temp1=FileLines(j+1,1)
            temp2=FileLines(j+1,2)
            temp3=FileLines(j+1,3)
            temp4=FileLines(j+1,4)
            FileLines(j+1,0)=FileLines(j,0)
            FileLines(j+1,1)=FileLines(j,1)
            FileLines(j+1,2)=FileLines(j,2)
            FileLines(j+1,3)=FileLines(j,3)
            FileLines(j+1,4)=FileLines(j,4)
            FileLines(j,0)=temp0
            FileLines(j,1)=temp1
            FileLines(j,2)=temp2
            FileLines(j,3)=temp3
            FileLines(j,4)=temp4
        end if
    next
next 

OutFile="Cemetery.html"
Set FileOut = FSO.CreateTextFile(OutFile,True)

FileOut.Write "<table style=""text-align: left; width: 650px; height: 60px;"" border='1' cellpadding='2' cellspacing='2'>" & vbCrLf
FileOut.Write "<tbody>" & vbCrLf
For l = 0 to i Step 1
  FileOut.Write "<tr><td>"&FileLines(l,0)&"</td><td>"&FileLines(l,3)&"</td><td>"&FileLines(l,4)&"</td></tr>"&vbCrLf
Next
FileOut.Write "<tbody>" & vbCrLf
FileOut.Close

for a = i - 1 To 0 Step -1
    for j = 0 to a
        if FileLines(j,3)>FileLines(j+1,3) then
            temp0=FileLines(j+1,0)
            temp1=FileLines(j+1,1)
            temp2=FileLines(j+1,2)
            temp3=FileLines(j+1,3)
            temp4=FileLines(j+1,4)
            FileLines(j+1,0)=FileLines(j,0)
            FileLines(j+1,1)=FileLines(j,1)
            FileLines(j+1,2)=FileLines(j,2)
            FileLines(j+1,3)=FileLines(j,3)
            FileLines(j+1,4)=FileLines(j,4)
            FileLines(j,0)=temp0
            FileLines(j,1)=temp1
            FileLines(j,2)=temp2
            FileLines(j,3)=temp3
            FileLines(j,4)=temp4
        end if
    next
next 

OutFile="First.html"
Set FileOut = FSO.CreateTextFile(OutFile,True)

FileOut.Write "<table style=""text-align: left; width: 650px; height: 60px;"" border='1' cellpadding='2' cellspacing='2'>" & vbCrLf
FileOut.Write "<tbody>" & vbCrLf
For l = 0 to i Step 1
  FileOut.Write "<tr><td>"&FileLines(l,0)&"</td><td>"&FileLines(l,3)&"</td><td>"&FileLines(l,4)&"</td></tr>"&vbCrLf
Next
FileOut.Write "<tbody>" & vbCrLf
FileOut.Close

for a = i - 1 To 0 Step -1
    for j = 0 to a
        if FileLines(j,4)>FileLines(j+1,4) then
            temp0=FileLines(j+1,0)
            temp1=FileLines(j+1,1)
            temp2=FileLines(j+1,2)
            temp3=FileLines(j+1,3)
            temp4=FileLines(j+1,4)
            FileLines(j+1,0)=FileLines(j,0)
            FileLines(j+1,1)=FileLines(j,1)
            FileLines(j+1,2)=FileLines(j,2)
            FileLines(j+1,3)=FileLines(j,3)
            FileLines(j+1,4)=FileLines(j,4)
            FileLines(j,0)=temp0
            FileLines(j,1)=temp1
            FileLines(j,2)=temp2
            FileLines(j,3)=temp3
            FileLines(j,4)=temp4
        end if
    next
next 

OutFile="Era.html"
Set FileOut = FSO.CreateTextFile(OutFile,True)

FileOut.Write "<table style=""text-align: left; width: 650px; height: 60px;"" border='1' cellpadding='2' cellspacing='2'>" & vbCrLf
FileOut.Write "<tbody>" & vbCrLf
For l = 0 to i Step 1
  FileOut.Write "<tr><td>"&FileLines(l,0)&"</td><td>"&FileLines(l,3)&"</td><td>"&FileLines(l,4)&"</td></tr>"&vbCrLf
Next
FileOut.Write "<tbody>" & vbCrLf
FileOut.Close

WScript.Echo "Done"
Wscript.Quit(0)

