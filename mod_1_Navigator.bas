Attribute VB_Name = "mod_1_Navigator"
Option Explicit

' =========================================================
' БЛОК 1: НАВИГАЦИЯ И ЗАКРЫТИЕ (Работает через PZ_SearchMain)
' =========================================================

Function CurrentOrder() As clsOrder
    Dim ord As New clsOrder
    ord.InitializeSearch ThisWorkbook.Sheets("PZ_Control").Range("PZ_SearchMain").Text
    Set CurrentOrder = ord
End Function

Sub PZ_Teleport()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim fVal As String: fVal = Trim(wsP.Range("PZ_SearchMain").Text)
    
    If fVal = "" Then MsgBox "Введите номер в поле поиска!", 48: Exit Sub
    
    UpdateSearchHistory fVal
    Update_PZ_WorkName fVal
    
    Dim ord As clsOrder: Set ord = CurrentOrder
    If ord.Rows.Count = 0 Then
        Application.StatusBar = "MES: Заказ " & fVal & " не найден"
        MsgBox "Заказ '" & fVal & "' не найден!", 48: Exit Sub
    End If
    
    Dim wsB As Worksheet: Set wsB = ord.BaseSheet
    
    ' --- ЛОГИКА ВЫБОРА УНИКАЛЬНОГО ЗАКАЗА (ТЕЛЕПОРТ) ---
    Dim colUnique As New Collection
    Dim rNum As Variant, fullName As String
    On Error Resume Next
    For Each rNum In ord.Rows
        fullName = Trim(wsB.Cells(Val(rNum), 15).Text)
        colUnique.Add fullName, CStr(fullName)
    Next rNum
    On Error GoTo 0
    
    Dim targetFullName As String
    If colUnique.Count = 1 Then
        targetFullName = colUnique(1)
    Else
        Dim promptStr As String
        promptStr = "Найдено несколько цехов для заказа " & fVal & ". Куда прыгаем?" & vbCrLf & vbCrLf
        Dim iChoice As Integer
        For iChoice = 1 To colUnique.Count
            promptStr = promptStr & iChoice & " - " & colUnique(iChoice) & vbCrLf
        Next iChoice
        promptStr = promptStr & vbCrLf & "Введите НОМЕР нужного варианта (1-" & colUnique.Count & "):"
        
        Dim choice As String
        choice = InputBox(promptStr, "MES: Уточнение адреса")
        
        If choice = "" Or Not IsNumeric(choice) Then
            Application.StatusBar = "MES: Телепорт отменен."
            Exit Sub
        End If
        If Val(choice) < 1 Or Val(choice) > colUnique.Count Then
            MsgBox "Неверный номер!", 16: Exit Sub
        End If
        targetFullName = colUnique(Val(choice))
        
        ' Фильтруем коллекцию строк, оставляя только выбранный заказ
        Dim filteredRows As New Collection
        For Each rNum In ord.Rows
            If Trim(wsB.Cells(Val(rNum), 15).Text) = targetFullName Then
                filteredRows.Add rNum
            End If
        Next rNum
        Set ord.Rows = filteredRows
    End If
    ' ---------------------------------------------------
    
    Dim idx As Long: idx = val(wsP.Range("PZ_TeleportIdx").Value) + 1
    If idx > ord.Rows.Count Then idx = 1
    wsP.Unprotect Password:=""
    wsP.Range("PZ_TeleportIdx").Value = idx
    wsP.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
    Dim targetRow As Long: targetRow = ord.Rows(idx)
    Set wsB = ord.BaseSheet
    
    Application.ScreenUpdating = True
    wsB.Parent.Activate: wsB.Activate
    On Error Resume Next: AppActivate wsB.Parent.Name: On Error GoTo 0
    DoEvents
    
    wsB.Cells(targetRow, 15).Select
    With ActiveWindow
        .ScrollRow = targetRow
        .SmallScroll Down:=1: .SmallScroll Up:=1
    End With
    
    Application.StatusBar = "MES Телепорт " & fVal & ": " & idx & " из " & ord.Rows.Count
    
     ' ДОПОЛНИТЕЛЬНЫЙ ПОИСК ПО ЗВР
    Dim zvrVal As String, zvrSearchTerm As String
    On Error Resume Next
    zvrVal = Trim(wsP.Range("PZ_SearchZVR").Text)
    On Error GoTo 0
    
    If zvrVal <> "" And zvrVal <> "Не найден" And zvrVal <> "Не найдена" And zvrVal <> fVal Then
        ' Если ЗВР не начинается с дефиса, добавляем его,
        ' чтобы IsMatch сделал поиск вхождения без строгих границ
        If Left(zvrVal, 1) <> "-" Then
            zvrSearchTerm = "-" & zvrVal
        Else
            zvrSearchTerm = zvrVal
        End If
        
        Dim zvrOrd As New clsOrder
        zvrOrd.InitializeSearch zvrSearchTerm
        
        If zvrOrd.Rows.Count > 0 Then
            Dim rowList As String
            Dim i As Long
            rowList = ""
            For i = 1 To zvrOrd.Rows.Count
                rowList = rowList & zvrOrd.Rows(i)
                If i < zvrOrd.Rows.Count Then rowList = rowList & ", "
            Next i
            MsgBox "По номеру ЗВР (" & zvrVal & ") найдены дополнительные строки в основной таблице." & vbCrLf & "Номера строк: " & rowList, vbInformation, "Дополнительный поиск по ЗВР"
        End If
    End If
    ResetFindDialog
End Sub

Sub PZ_ProcessRow()
    CurrentOrder.ApplyStyling
    MsgBox "Готово!", 64
End Sub

' МОСТЫ: Добавить участок в СУЩЕСТВУЮЩИЙ заказ (Закрытие)
Sub Add_KSU(): CurrentOrder.AddSection "КСУ АК", "Работа КСУ": End Sub
Sub Add_SU():  CurrentOrder.AddSection "СУ АК", "Работа СУ":   End Sub
Sub Add_CNC(): CurrentOrder.AddSection "Группа ЧПУ", "Работа ЧПУ": End Sub

' =========================================================
' БЛОК 2: ХИРУРГИЧЕСКАЯ ИСТОРИЯ (Без сдвига ячеек)
' =========================================================
Sub UpdateSearchHistory(ByVal newVal As String)
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim histRange As Range: Set histRange = wsP.Range("PZ_SearchHistory")
    Dim i As Integer
    
    If newVal = "" Or newVal = "Не найден" Or newVal = "Не найдена" Then Exit Sub
    
    Application.EnableEvents = False
    wsP.Unprotect Password:=""
    
    Dim mIdx As Variant
    mIdx = Application.Match(newVal, histRange, 0)
    
    If Not IsError(mIdx) Then
        For i = mIdx To 2 Step -1 ' Изменено для работы с Range напрямую, если он отвязан от колонок. Но оставим логику сдвига через Cells, если она жестко привязана.
            histRange.Cells(i, 1).Value = histRange.Cells(i - 1, 1).Value
        Next i
    Else
        For i = 10 To 2 Step -1
            histRange.Cells(i, 1).Value = histRange.Cells(i - 1, 1).Value
        Next i
    End If
    
    histRange.Cells(1, 1).Value = newVal
    wsP.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    Application.EnableEvents = True
End Sub

' --- ВОССТАНОВЛЕНИЕ НАСТРОЕК ПОИСКА (Ctrl+F) ---
Sub ResetFindDialog()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Dim dummy As Range
    ' Делаем пустой поиск с параметром xlPart, чтобы сбросить "Ячейка целиком"
    On Error Resume Next
    Set dummy = ws.Cells.Find(What:="", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)
    On Error GoTo 0
End Sub


' --- ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: ПОИСК НАИМЕНОВАНИЯ РАБОТЫ В tblZVR_Master ---
Sub Update_PZ_WorkName(ByVal sSearch As String)
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim tbl As ListObject, ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects("tblZVR_Master")
        On Error GoTo 0
        If Not tbl Is Nothing Then Exit For
    Next ws
    
    If tbl Is Nothing Then Exit Sub
    
    Dim foundRow As Range
    ' Ищем по Заказу (столбец 2)
    Set foundRow = tbl.ListColumns(2).DataBodyRange.Find(What:=sSearch, LookIn:=xlValues, LookAt:=xlWhole)
    If foundRow Is Nothing And IsNumeric(sSearch) Then
        Set foundRow = tbl.ListColumns(2).DataBodyRange.Find(What:=CLng(sSearch), LookIn:=xlValues, LookAt:=xlWhole)
    End If
    
    If Not foundRow Is Nothing Then
        wsP.Unprotect Password:=""
        wsP.Range("PZ_WorkName").Value = foundRow.Offset(0, 2).Value
        wsP.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
        Exit Sub
    End If
    
    ' Ищем по ЗВР (столбец 1)
    Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(What:=sSearch, LookIn:=xlValues, LookAt:=xlWhole)
    If foundRow Is Nothing And IsNumeric(sSearch) Then
        Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(What:=CLng(sSearch), LookIn:=xlValues, LookAt:=xlWhole)
    End If
    
    If Not foundRow Is Nothing Then
        wsP.Unprotect Password:=""
        wsP.Range("PZ_WorkName").Value = foundRow.Offset(0, 3).Value
        wsP.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    Else
        wsP.Unprotect Password:=""
        wsP.Range("PZ_WorkName").ClearContents
        wsP.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub
