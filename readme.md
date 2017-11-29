# excel
Excel tools and workbooks

- File > Options
- Add-Ins > Manage: Excel Add-ins "Go" > Browse > Select vbaTools.xlsm
- Trust Center > "Trust Center Settings" Macro Settings > "Check" Traust access to the VBA project object model
- Add the following code to "ThisWorkbook" for automatic backup and restore on save

```
Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Application.Run ("vbaTools.xlsm!VBABackup")
End Sub

Sub Workbook_Open()
Application.Run ("vbaTools.xlsm!VBARestore")
End Sub
```
