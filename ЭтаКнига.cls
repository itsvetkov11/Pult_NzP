VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' Включаем спец-режим защиты для макросов
    ' Это позволит Clear_Pulse и Бэкапам работать на закрытом листе
    On Error Resume Next
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        sh.Protect Password:="", UserInterfaceOnly:=True, _
                   AllowFiltering:=True, AllowSorting:=True
    Next sh
    
    ' Запуск логики бэкапа при открытии
    Run_Smart_Backup_Logic
End Sub

