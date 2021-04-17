Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub exchangeModule()

'definovanie premennych
Dim wb As Workbook
Dim ws As Worksheet

Set fso = CreateObject("Scripting.FileSystemObject")
Set Data = Workbooks("import module.xlsm")

'Input okno pre adresu k priecinku
cesta = InputBox("Zadaj cestu k objektu")
Set fldr = fso.GetFolder(cesta)
'Vypnutie screen updating, kvoli preblikavaniu obrazovky
Application.ScreenUpdating = False

'Loop pre kazdy subor v priecinku
For Each wbFile In fldr.Files
    
'Podmienka pre otvorenie iba .xlsm suborov obsahujucich "CF" v nazve
    If fso.GetExtensionName(wbFile.Name) = "xlsm" And InStr(wbFile.Name, "CF") Then

'Otvorenie dokumentu
        Set wb = Workbooks.Open(wbFile.Path, ReadOnly:=False)
        Set ws = wb.Sheets(1)
            'Vymazanie povodnych modulov
            For Each m In wb.VBProject.vbcomponents
                If m.Name = "Module7" Or m.Name = "Module8" Then
                    wb.VBProject.vbcomponents.Remove m
                End If
            'Import spravnych modulov
            Next
                wb.VBProject.vbcomponents.Import ("C:\Users\rypak.QUALIFORM\Desktop\Module7.bas")
                wb.VBProject.vbcomponents.Import ("C:\Users\rypak.QUALIFORM\Desktop\Module8.bas")
                
          
            
                
'Ulozenie a zatvorenie dokumentu
wb.Close True
End If

'dalsi dokument
Next wbFile
Application.ScreenUpdating = True
End Sub

