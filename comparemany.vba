Option Explicit

Sub CreateReport_Click()
Main
End Sub
Sub Main()
On Error Resume Next
Dim i, j, k As Integer
Dim wb As Workbook

Dim cn As Object, cn_rec, cn_rec2 As Object, FilePath, Output As String
Dim LastRowSource, LastRowDest, LastColumn As Long
Dim SourceSheet, SourceTable, SourceCol, WhereCondi As String
Dim objNewWorkbook As Workbook

Dim objNewWorksheet As Worksheet
 Set wb = Workbooks("CompareDataV2.xlsm")
   SourceSheet = wb.Sheets("GUI").Range("L7")
   SourceTable = wb.Sheets("GUI").Range("Q7")
   SourceCol = wb.Sheets("GUI").Range("F7")
   WhereCondi = ""
If wb.Sheets("GUI").Range("T7") <> "" Then
 WhereCondi = "Where" & " " & wb.Sheets("GUI").Range("T7")
End If
    
If SourceSheet = "" Or SourceTable = "" Or SourceCol = "" Then
MsgBox "Please Entered value"
Exit Sub
End If
     
    Set objNewWorkbook = Excel.Application.Workbooks.Add
    Set objNewWorksheet = objNewWorkbook.Sheets(1)

    LastRowSource = wb.Sheets("FilePath").Cells(Rows.Count, "A").End(xlUp).Row
 
    For j = 2 To LastRowSource
    LastRowDest = objNewWorkbook.Worksheets("Sheet1").Cells(Rows.Count, "B").End(xlUp).Row
    For k = 7 To 19
     objNewWorkbook.Sheets("Sheet1").Range("A" & LastRowDest + 1) = wb.Sheets("FilePath").Range("A" & j)
        'Debug.Print wb.Sheets("Sheet2").Range("G" & j)
        Set cn = CreateObject("ADODB.Connection")
        Set cn_rec2 = CreateObject("ADODB.Recordset")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .ConnectionString = "Data Source=" & wb.Sheets("FilePath").Range("B" & j) & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        .Open
    End With
     
  
    LastColumn = objNewWorkbook.Worksheets("Sheet1").Cells(LastRowDest + 1, Columns.Count).End(xlToLeft).Column
    'Debug.Print wb.Sheets("GUI").Range("L" & k)
    cn_rec = "select " & SourceCol & " from " & "[" & wb.Sheets("GUI").Range("L" & k) & SourceTable & "] " & WhereCondi
  Debug.Print cn_rec
    cn_rec2.Open cn_rec, cn
            With cn_rec2.objRecordset
            objNewWorkbook.Sheets("Sheet1").Cells(LastRowDest + 1, LastColumn + 1).Value = wb.Sheets("GUI").Range("L" & k)
            For i = 0 To cn_rec2.Fields.Count - 1
                objNewWorkbook.Sheets("Sheet1").Cells(LastRowDest + 2, i + LastColumn + 1).Value = cn_rec2.Fields(i).Name
                
            Next i
                objNewWorkbook.Sheets("Sheet1").Cells(LastRowDest + 2, LastColumn + 1).CopyFromRecordset cn_rec2
                With ActiveWorkbook.Styles("Normal").Font
                 .Size = 10
               End With
            End With
       ' cn.Close
        Set cn_rec = Nothing
        Set cn_rec2 = Nothing
        Next k
        
      Next j
    
End Sub
