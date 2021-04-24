' 1. Get clipboard text
Set objHtml = CreateObject("htmlfile")
clipBoard = objHtml.ParentWindow.ClipboardData.GetData("text")
If clipBoard <> "" Then ' Do nothing if empty

Dim clipBoardText
clipBoardText = CStr(clipBoard)
clipBoardText = Trim(clipBoardText)
clipBoardText = LCase(clipBoardText)
splitText = Split(clipBoardText, " ")
textToSpeak = splitText(0) ' get only first word

' Test if it has chinese character to avoid error
Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "[^a-zA-Z0-9\s]"
doesWriteLog = True
If objRegExp.Test(clipBoardText) Then 
    textToSpeak = "character mistake"
    doesWriteLog = false
End If

' 2. Speak it
Set sapi = CreateObject("sapi.spvoice")
If Len(textToSpeak) >= 30 Then 
    sapi.Speak "text too long"
    doesWriteLog = false
ElseIf Len(textToSpeak) < 30 Then 
    sapi.Speak textToSpeak
End If

' 3. Save a log
Set fso = CreateObject("Scripting.FileSystemObject")
strFile = "words.txt"

' Read old
strLine = ""
If (fso.FileExists(strFile)) Then
    Set objFileRead = fso.OpenTextFile(strFile)
    Do Until objFileRead.AtEndOfStream
        strLine = objFileRead.ReadLine
    Loop
    objFileRead.Close
End If

' Write new
Set objFileWrite = fso.CreateTextFile(strFile, True)
If doesWriteLog Then
    objFileWrite.Write strLine + textToSpeak + " " & vbCrLf
Else
    objFileWrite.Write strLine + " " & vbCrLf
End If
objFileWrite.Close
End If ' If clipboard empty