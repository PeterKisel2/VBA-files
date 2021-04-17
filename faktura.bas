Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub getDataFromWbs()

Dim wb As Workbook
Dim ws As Worksheet

Set fso = CreateObject("Scripting.FileSystemObject")
Set Data = Workbooks("fakt�ra.xlsm")

'This is where you put YOUR folder name
cesta = InputBox("Zadaj cestu k objektu")
Set fldr = fso.GetFolder(cesta)
Application.ScreenUpdating = False

'Next available Row on Master Workbook
y = Data.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row + 1

'Loop through each file in that folder
For Each wbFile In fldr.Files
    y = Data.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row + 1
'Make sure looping only through files ending in .xlsx (Excel files)
    If fso.GetExtensionName(wbFile.Name) = "xlsm" Then

'Open current book
        Set wb = Workbooks.Open(wbFile.Path, ReadOnly:=True)
        Set ws = wb.Sheets(1)

'load data
            On Error Resume Next
            Data.Sheets("data").Cells(y, 1).Value = ActiveWorkbook.Name
            Data.Sheets("data").Cells(y, 2).Value = ws.Range("id_datum_vyhotovenia").Value
            Data.Sheets("data").Cells(y, 3).Value = ws.Range("id_cislo").Value
                If CBool(InStr(ActiveWorkbook.Name, "CF")) Then
                    'Konzistencia
                    Data.Sheets("data").Cells(y, 7).Value = Application.WorksheetFunction.Count(ws.Range("O1:O27"))
                    'Objemovka
                    Data.Sheets("data").Cells(y, 8).Value = Application.WorksheetFunction.Count(ws.Range("R1:R27"))
                    'Vzduch
                    Data.Sheets("data").Cells(y, 9).Value = Application.WorksheetFunction.Count(ws.Range("Q1:Q27"))
                End If
                
                If CBool(InStr(ActiveWorkbook.Name, "CR")) Then
                    'Tlak
                    Data.Sheets("data").Cells(y, 10).Value = Application.WorksheetFunction.Count(ws.Range("C50:C52"))
                End If
                
                If CBool(InStr(ActiveWorkbook.Name, "DPW")) Then
                    'Vodotes
                    Data.Sheets("data").Cells(y, 11).Value = 1
                End If
                
                If CBool(InStr(ActiveWorkbook.Name, "WA")) Then
                    'Nasiak
                    Data.Sheets("data").Cells(y, 12).Value = 1
                End If
                
                If CBool(InStr(ActiveWorkbook.Name, "CHD")) Then
                    'Soli
                    If ws.Range("H56") = 50 Then
                    Data.Sheets("data").Cells(y, 13).Value = 1
                    End If
                    If ws.Range("H56") = 100 Then
                    Data.Sheets("data").Cells(y, 14).Value = 1
                    End If
                End If
                
                If CBool(InStr(ActiveWorkbook.Name, "FR")) Then
                    'Mrazy
                    If ws.Range("id_cycle") = 25 Then
                    Data.Sheets("data").Cells(y, 15).Value = 1
                    End If
                    If ws.Range("id_cycle") = 50 Then
                    Data.Sheets("data").Cells(y, 16).Value = 1
                    End If
                    If ws.Range("id_cycle") = 100 Then
                    Data.Sheets("data").Cells(y, 17).Value = 1
                    End If
                    If ws.Range("id_cycle") = 150 Then
                    Data.Sheets("data").Cells(y, 18).Value = 1
                    End If
                End If
                
                If CBool(InStr(ActiveWorkbook.Name, "FS")) Then
                    'ohyb CB
                    Data.Sheets("data").Cells(y, 2).Value = ws.Range("N50").Value
                    Data.Sheets("data").Cells(y, 19).Value = 1
                End If
                
                If CBool(InStr(ActiveWorkbook.Name, "SS")) Or CBool(InStr(ActiveWorkbook.Name, "TS")) Then
                    'ohyb CB
                    Data.Sheets("data").Cells(y, 20).Value = 1
                End If

            y = y + 1

'Close current book
wb.Close False
End If

'go to next file
Next wbFile

'filter dates
ActiveSheet.Range("$A$1:$AJ$1583").AutoFilter Field:=2, Criteria1:= _
        ">=" & "2021-01-16", Operator:=xlAnd, Criteria2:="<=" & "2021-02-15"
        
Application.ScreenUpdating = True
End Sub

