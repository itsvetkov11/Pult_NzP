import re

def fix_navigator():
    with open('mod_1_Navigator.bas', 'r', encoding='utf-8') as f:
        content = f.read()

    search_block = """    Dim ord As clsOrder: Set ord = CurrentOrder
    If ord.Rows.Count = 0 Then
        Application.StatusBar = "MES: Заказ " & fVal & " не найден"
        MsgBox "Заказ '" & fVal & "' не найден!", 48: Exit Sub
    End If"""

    replace_block = """    Dim ord As clsOrder: Set ord = CurrentOrder
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
    ' ---------------------------------------------------"""

    if search_block in content:
        content = content.replace(search_block, replace_block)
        with open('mod_1_Navigator.bas', 'w', encoding='utf-8') as f:
            f.write(content)
        print("Updated mod_1_Navigator.bas")
    else:
        print("Could not find block in mod_1_Navigator.bas")

def fix_clsOrder():
    with open('clsOrder.cls', 'r', encoding='utf-8') as f:
        content = f.read()

    search_block = """    Me.GatherRows
    
    Dim colUnique As New Collection
    Dim rNum As Variant, fullName As String
    On Error Resume Next
    For Each rNum In Rows
        fullName = Trim(wsB.Cells(val(rNum), 15).Text)
        colUnique.Add fullName, CStr(fullName)
    Next rNum
    On Error GoTo 0
    
    For Each rNum In Rows"""

    replace_block = """    Me.GatherRows
    
    Dim colUnique As New Collection
    Dim rNum As Variant, fullName As String
    On Error Resume Next
    For Each rNum In Rows
        fullName = Trim(wsB.Cells(Val(rNum), 15).Text)
        colUnique.Add fullName, CStr(fullName)
    Next rNum
    On Error GoTo 0
    
    Dim targetFullName As String
    If colUnique.Count = 1 Then
        targetFullName = colUnique(1)
    Else
        Dim promptStr As String
        promptStr = "Найдено несколько цехов для заказа. Что красим?" & vbCrLf & vbCrLf
        Dim idx As Integer
        For idx = 1 To colUnique.Count
            promptStr = promptStr & idx & " - " & colUnique(idx) & vbCrLf
        Next idx
        promptStr = promptStr & vbCrLf & "Введите НОМЕР нужного варианта (1-" & colUnique.Count & "):"
        
        Dim choice As String
        choice = InputBox(promptStr, "MES: Уточнение адреса")
        
        If choice = "" Or Not IsNumeric(choice) Then
            Application.StatusBar = "MES: Покраска отменена."
            Exit Sub
        End If
        If Val(choice) < 1 Or Val(choice) > colUnique.Count Then
            MsgBox "Неверный номер!", 16: Exit Sub
        End If
        targetFullName = colUnique(Val(choice))
        
        ' Фильтруем коллекцию строк, оставляя только выбранный заказ
        Dim filteredRows As New Collection
        For Each rNum In Rows
            If Trim(wsB.Cells(Val(rNum), 15).Text) = targetFullName Then
                filteredRows.Add rNum
            End If
        Next rNum
        Set Rows = filteredRows
        
        ' Обновляем colUnique, чтобы логика рамок ниже работала только для одной группы
        Set colUnique = New Collection
        colUnique.Add targetFullName
    End If
    
    For Each rNum In Rows"""

    if search_block in content:
        content = content.replace(search_block, replace_block)
        with open('clsOrder.cls', 'w', encoding='utf-8') as f:
            f.write(content)
        print("Updated clsOrder.cls")
    else:
        print("Could not find block in clsOrder.cls")


if __name__ == '__main__':
    fix_navigator()
    fix_clsOrder()

