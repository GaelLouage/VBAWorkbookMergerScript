
Imports System.Net.Mime.MediaTypeNames

Set fd = Application.FileDialog(msoFileDialogFilePicker)

'get the number of the button chosen

Dim FileChosen As Integer

FileChosen = fd.Show

If FileChosen <> -1 Then

MsgBox "Cancelled"
' add a stop to all if cancelled'
Else
'display name and path of file chosen
GetFilePath = fd.SelectedItems(1)
End If

End Function

Function AddSheet() As Worksheet
    Dim secondSheet As Worksheet
    ' Add a new worksheet and name it "Sheet2"
    Set secondSheet = Sheets.Add
    secondSheet.Name = "Sheet2"
    
    ' Return the created worksheet
    Set AddSheet = secondSheet
End Function

Sub CopyDataBetweenWorkbooks()
    Dim sourceWorkbook As Workbook
    Dim sourceWorkbookTwo As Workbook
    Dim targetWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceSheet2 As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceFilePath As String
    Dim SourceFilePathSecond As String
    Dim targetFilePath As String
    AddSheet()
    ' Define file paths
    sourceFilePath = GetFilePath()
    SourceFilePathSecond = GetFilePath()
    targetFilePath = "C:\Users\louag\OneDrive\Bureau\Book1.xlsm"

    ' Open the source workbook
    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Set sourceSheet = sourceWorkbook.Sheets("Sheet1") ' Change to the correct sheet name
    
     ' Open the second source workbook
    Set sourceWorkbookTwo = Workbooks.Open(SourceFilePathSecond)
    Set sourceSheet2 = sourceWorkbookTwo.Sheets("Sheet1") ' Change to the correct sheet name
    
    ' Open the target workbook
    Set targetWorkbook = Workbooks.Open(targetFilePath)
    Set targetSheet = targetWorkbook.Sheets("Sheet1") ' Change to the correct sheet name

  
    ' Copy data (e.g., A1:D100 from source to A1 in target)
    sourceSheet.Range("A1:D100").Copy
    targetSheet.Range("A1").PasteSpecial Paste:=xlPasteValues ' Paste values only



    Set targetWorkbook = Workbooks.Open(targetFilePath)
    Set targetSheet = targetWorkbook.Sheets("Sheet2") ' Change to the correct sheet name
    
        ' Copy data (e.g., A1:D100 from source to A1 in target)
    sourceSheet2.Range("A1:D100").Copy
    targetSheet.Range("A1").PasteSpecial Paste:=xlPasteValues ' Paste values only

    ' Save and close workbooks
    targetWorkbook.Save
    sourceWorkbook.Close False ' Close source without saving changes
    sourceWorkbookTwo.Close True ' Save and close target
    ' Clean up
    Application.CutCopyMode = False
    Set sourceSheet = Nothing
    Set targetSheet = Nothing
    Set sourceWorkbook = Nothing
    Set targetWorkbook = Nothing
    Set sourceSheet2 = Nothing
    Set sourceWorkbookTwo = Nothing
    MsgBox "Data copied successfully!", vbInformation
End Sub