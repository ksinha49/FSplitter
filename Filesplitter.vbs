Dim  Counter

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")

InputFile=WScript.Arguments.Unnamed(0)
OutputFile=WScript.Arguments.Unnamed(1)
RecordSize=160000


Set objTextFile = objFSO.OpenTextFile (InputFile, ForReading)

Counter = 0
FileCounter = 0
Set objOutTextFile = Nothing
 
Do Until objTextFile.AtEndOfStream
  
    if Counter = 0 Or Counter = RecordSize Then
        Counter = 0
        FileCounter = FileCounter + 1
        if Not objOutTextFile is Nothing then objOutTextFile.Close
        Set objOutTextFile = objFSO.OpenTextFile( OutputFile & "_" & FileCounter & ".csv", ForWriting, True)
    end if
    strNextLine = objTextFile.Readline
    objOutTextFile.WriteLine(strNextLine)
    Counter = Counter + 1

Loop
objTextFile.Close
objOutTextFile.Close