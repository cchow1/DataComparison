Option Explicit

Sub SearchBt_Click()
On Error Resume Next
Dim wb As Workbook
Dim strFileToOpen As String
Set wb = Workbooks("CompareData.xlsm")
strFileToOpen = Application.GetOpenFilename(FileFilter:="(*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Open Database File", MultiSelect:=False)
 Workbooks.Open Filename:=strFileToOpen

wb.Sheets("GUI").Range("G18") = strFileToOpen
ReportDisplay1
End Sub

Sub ReportDisplay1()
'Dim tclsconn As New clsconn
On Error Resume Next
Dim i, j As Integer
Dim wb As Workbook
Dim cn As Object, cn_rec, cn_rec2 As Object, FilePath, Output As String
Dim LastRowSource, LastRowDest As Long
Dim SourceSheet, SourceTable, SourceCol, WhereCondi As String
Dim strConnString As String
Dim clsSQL As New clsconn
Dim connectionObject As ADODB.Connection
'Set connectionObject = ADODB.Connection
Set wb = Workbooks("CompareData.xlsm")
 clsSQL.ConnProvider = "Microsoft.ACE.OLEDB.12.0"
 clsSQL.ConnProperties = "Excel 12.0; HDR=YES"
 strConnString = wb.Sheets("GUI").Range("G18") 'I take it this named range holds a filepath?
clsSQL.ConnString = strConnString
Set connectionObject = clsSQL.GetConnectionObject

Dim objNewWorkbook As Workbook
Dim objNewWorksheet As Worksheet
    
'If SourceSheet = "" Or SourceTable = "" Or SourceCol = "" Then
   ' MsgBox "Please Entered value"
   ' Exit Sub
'End If
   
    Set objNewWorkbook = Excel.Application.Workbooks.Add
      Set objNewWorksheet = objNewWorkbook.Sheets(1)
Dim RowCount As Integer
 
 
     RowCount = WorksheetFunction.CountA(wb.Sheets("GUI").Range("F21:F27"))
   For j = 0 To RowCount
   Set cn_rec2 = CreateObject("ADODB.Recordset")

    SourceSheet = wb.Sheets("GUI").Range("L" & 21 + j)
    SourceTable = wb.Sheets("GUI").Range("Q" & 21 + j)
    SourceCol = wb.Sheets("GUI").Range("F" & 21 + j)
    WhereCondi = ""
If wb.Sheets("GUI").Range("T" & 21 + j) <> "" Then
    WhereCondi = "Where" & " " & wb.Sheets("GUI").Range("T" & 21 + j)
End If
    cn_rec = "select" & SourceCol & "from" & "[" & SourceSheet & SourceTable & "] " & WhereCondi
    Debug.Print cn_rec
    cn_rec2.Open cn_rec, connectionObject
    LastRowDest = objNewWorkbook.Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
        With cn_rec2.objRecordset
            For i = 0 To cn_rec2.Fields.Count - 1
                objNewWorkbook.Sheets("Sheet1").Cells(LastRowDest + 2, i).Value = cn_rec2.Fields(i).Name
            Next i
          LastRowDest = objNewWorkbook.Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
                'Debug.Print LastRowSource; wb.Sheets("Sheet2").Range("A" & j)
                objNewWorkbook.Sheets("Sheet1").Range("A" & LastRowDest + 1).CopyFromRecordset cn_rec2
                With ActiveWorkbook.Styles("Normal").Font
                 .Size = 10
               End With
        End With
       
       
      j = j + 1
      Next j

End Sub
