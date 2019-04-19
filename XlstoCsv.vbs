if WScript.Arguments.Count < 2 Then
    WScript.Echo "Please specify the source and the destination files. Usage: ExcelToCsv <xls/xlsx source file> <csv destination file>"
    Wscript.Quit
End If

csv_format = 6

Set objFSO = CreateObject("Scripting.FileSystemObject")

src_file = objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
dest_file = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))

Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)

oBook.SaveAs dest_file, csv_format

oBook.Close False
oExcel.Quit

Const FOR_READING = 1 
Const FOR_WRITING = 2 
strFileName = WScript.Arguments.Unnamed(1)
iNumberOfLinesToDelete = 1 
 
Set objFS = CreateObject("Scripting.FileSystemObject") 
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING) 
strContents = objTS.ReadAll 
objTS.Close 
 
arrLines = Split(strContents, vbNewLine) 
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING) 
 
For i=0 To UBound(arrLines) 
   If i > (iNumberOfLinesToDelete - 1) Then 
      objTS.WriteLine arrLines(i) 
   End If 
Next 