Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub getDataFromWbs()

Dim wb As Workbook
Dim ws As Worksheet
Dim find_rng As Range

Set fso = CreateObject("Scripting.FileSystemObject")
Set Data = Workbooks("moduly pruznosti.xlsm")

'This is where you put YOUR folder name
cesta = InputBox("Zadaj cestu k objektu")
Set fldr = fso.GetFolder(cesta)

Application.ScreenUpdating = False

'Next available Row on Master Workbook
y = Data.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row + 1

'Loop through each file in that folder
For Each wbFile In fldr.Files
    
'Make sure looping only through files ending in .xlsx (Excel files)
    If fso.GetExtensionName(wbFile.Name) = "xlsm" And CBool(InStr(wbFile.Name, "CF")) Then

'Open current book
        Set wb = Workbooks.Open(wbFile.Path, ReadOnly:=True)
        Set ws = wb.Sheets(1)
        Set find_rng = ws.Range("L26:S42").Find("100x100x400 mm 4 ks/ pcs  sk��ka modulu pru�nosti / modulus of elasticity test")
                    If Not find_rng Is Nothing Then
                        
                    On Error Resume Next
                    Data.Sheets("data").Cells(y, 2).Value = wb.Name
                    Data.Sheets("data").Cells(y, 3).Value = ws.Range("id_datum_zhotovenia").Value
                    'Data.Sheets("data").Cells(y, 4).Value = ws.Range("id_beton").Value
                    'Data.Sheets("data").Cells(y, 6).Value = ws.Range("id_datum_zhotovenia").Value
                    'Data.Sheets("data").Cells(y, 10).Value = ws.Range("D20").Value
                    Data.Sheets("data").Cells(y, 4).Value = ws.Range("id_konstrukcia").Value
                    y = y + 1
                    'Konzistencia
                    'Data.Sheets("data").Cells(y, 7).Value = ws.Range("M50").Value
                    'Objemovka
                    'Data.Sheets("data").Cells(y, 8).Value = ws.Range("M51").Value
                    'Vzduch
                    End If
                    
            

'Close current book
    wb.Close False
    End If

'Next file
Next wbFile
Application.ScreenUpdating = True
End Sub

