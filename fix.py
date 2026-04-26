import sys

def modify_file(filepath):
    with open(filepath, 'rb') as f:
        content = f.read().decode('cp1251')
    
    # 1. Protect -> Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    content = content.replace('wsP.Protect\r\n', 'wsP.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True\r\n')
    content = content.replace('wsP.Protect\n', 'wsP.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True\n')
    
    content = content.replace('ws.Protect\r\n', 'ws.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True\r\n')
    content = content.replace('ws.Protect\n', 'ws.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True\n')
    
    # Unprotect
    content = content.replace('wsP.Unprotect\r\n', 'wsP.Unprotect Password:=""\r\n')
    content = content.replace('wsP.Unprotect\n', 'wsP.Unprotect Password:=""\n')

    content = content.replace('ws.Unprotect\r\n', 'ws.Unprotect Password:=""\r\n')
    content = content.replace('ws.Unprotect\n', 'ws.Unprotect Password:=""\n')

    with open(filepath, 'wb') as f:
        f.write(content.encode('cp1251'))

for file in ["mod_1_Navigator.bas", "mod_2_Constructor.bas", "mod_3_Archive_Utility.bas", "clsOrder.cls"]:
    modify_file(file)

def modify_file2(filepath):
    with open(filepath, 'rb') as f:
        content = f.read().decode('cp1251')

    if filepath == "mod_1_Navigator.bas":
        parts = content.split('End Sub\r\n')
        # find where PZ_Teleport ends
        for i, p in enumerate(parts):
            if 'PZ_Teleport' in p and not 'Sub PZ_Teleport' in parts[i+1]:
                parts[i] = p + '    ResetFindDialog\r\n'
                break
        content = 'End Sub\r\n'.join(parts)
        
        reset_find = "\r\n' --- ВОССТАНОВЛЕНИЕ НАСТРОЕК ПОИСКА (Ctrl+F) ---\r\nSub ResetFindDialog()\r\n    Dim ws As Worksheet: Set ws = ActiveSheet\r\n    If ws Is Nothing Then Exit Sub\r\n    Dim dummy As Range\r\n    ' Делаем пустой поиск с параметром xlPart, чтобы сбросить \"Ячейка целиком\"\r\n    On Error Resume Next\r\n    Set dummy = ws.Cells.Find(What:=\"\", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)\r\n    On Error GoTo 0\r\nEnd Sub\r\n"
        content += reset_find
        
    elif filepath == "Лист1.cls":
        content = content.replace('    tbl.DataBodyRange.Find What:="", LookAt:=xlPart\r\n', '    ResetFindDialog\r\n')
        content = content.replace('    tbl.DataBodyRange.Find What:="", LookAt:=xlPart\n', '    ResetFindDialog\n')
        
    elif filepath == "mod_2_Constructor.bas":
        parts = content.split('End Sub\r\n')
        # find where Show_Ch_Hint ends
        for i, p in enumerate(parts):
            if 'Sub Show_Ch_Hint' in p:
                parts[i] = p + '    ResetFindDialog\r\n'
                break
        content = 'End Sub\r\n'.join(parts)

    with open(filepath, 'wb') as f:
        f.write(content.encode('cp1251'))

for file in ["mod_1_Navigator.bas", "mod_2_Constructor.bas", "Лист1.cls"]:
    modify_file2(file)
