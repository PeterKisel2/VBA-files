Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub getDataFromWbs()

Dim wb As Workbook
Dim ws As Worksheet

Set fso = CreateObject("Scripting.FileSystemObject")

'This is where you put YOUR folder name
cesta = InputBox("Zadaj cestu k objektu")
Set fldr = fso.GetFolder(cesta)
Application.ScreenUpdating = False

'Loop through each file in that folder
For Each wbFile In fldr.Files

'Make sure looping only through files ending in .xlsm (Excel files) and contain specific string
    If fso.GetExtensionName(wbFile.Name) = "xlsm" And CBool(InStr(wbFile.Name, "CHD")) Then
    
    'Open current book
        Set wb = Workbooks.Open(wbFile.Path, ReadOnly:=False)
        Set ws = wb.Sheets(1)
    
            On Error Resume Next
            'change desired value
            ws.Range("AE11").Value = 2020

            'enable stamp and sign
            ActiveSheet.Shapes("razitko").Visible = True
            ActiveSheet.Shapes("podpis_Varga_Jozo").Visible = True
            
            'run pdf macro (button)
            wb.Application.Run ActiveSheet.Shapes("tla��tko 157").OnAction
            
            'disable stamp and sign
            ActiveSheet.Shapes("razitko").Visible = False
            ActiveSheet.Shapes("podpis_Varga_Jozo").Visible = False

            wb.Close True
            
    ElseIf fso.GetExtensionName(wbFile.Name) = "xlsm" And CBool(InStr(wbFile.Name, "FR")) Then
    
    'Open current book
        Set wb = Workbooks.Open(wbFile.Path, ReadOnly:=False)
        Set ws = wb.Sheets(1)
    
            On Error Resume Next
            ws.Range("AE9").Value = 2020
    
            ActiveSheet.Shapes("razitko").Visible = True
            ActiveSheet.Shapes("podpis_Varga_Jozo").Visible = True
            
            wb.Application.Run ActiveSheet.Shapes("Button 180").OnAction
            
            ActiveSheet.Shapes("razitko").Visible = False
            ActiveSheet.Shapes("podpis_Varga_Jozo").Visible = False
            wb.Close True
'Close current book

    End If

Next wbFile
Application.ScreenUpdating = True
End Sub


