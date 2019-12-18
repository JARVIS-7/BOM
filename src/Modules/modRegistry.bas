Attribute VB_Name = "modRegistry"
Option Explicit

Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_USERS = &H80000003
'Public Const HKEY_PERFORMANCE_DATA = &H80000004 '(nur NT)
'Public Const HKEY_CURRENT_CONFIG = vbWindowBackground
'Public Const HKEY_DYN_DATA = &H80000006

' intern
Private Const KEY_QUERY_VALUE As Long = &H1
'Private Const KEY_SET_VALUE As Long = &H2
'Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
'Private Const KEY_CREATE_LINK As Long = &H20
Private Const KEY_READ As Long = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
'Const KEY_WRITE = KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
'Const KEY_EXECUTE = KEY_READ
'Private Const KEY_ALL_ACCESS As Long = KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK
Private Const ERROR_SUCCESS As Long = 0&

'Const REG_NONE = 0      ' No value type
Private Const REG_SZ As Long = 1&       ' Unicode nul terminated string

'Const REG_EXPAND_SZ = 2 ' Unicode nul terminated string (with environment variable references)
'Const REG_BINARY = 3    ' Free form binary
Private Const REG_DWORD As Long = 4    ' 32-bit number

'Const REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
'Const REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
'Const REG_LINK = 6                ' Symbolic Link (unicode)
'Const REG_MULTI_SZ = 7            ' Multiple Unicode strings

'Private Const REG_OPTION_NON_VOLATILE As Long = &H0
'Private Const REG_CREATED_NEW_KEY As Long = &H1

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
'Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'Private Declare Function RegSetValueEx_DWord Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
'Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Any) As Long


' Prüft auf das Vorhandensein eines Schlüssels.
' Diese Funktion sollten Sie aufrufen bevor Sie einen neuen Eintrag hinzufügen

Private Function ExistKey(lRoot&, sSchluessel$) As Boolean
    
    'MD-Marker , Function wird nicht aufgerufen
    
    ' Root ist entweder HKEY_CURRENT_USER oder HKEY_LOCAL_MACHINE
    'Dim lResult&, keyhandle&
    '
    'lResult = RegOpenKeyEx(Root, schlüssel, 0, KEY_READ, keyhandle)
    'If lResult = ERROR_SUCCESS Then RegCloseKey keyhandle
    'ExistKey = (lResult = ERROR_SUCCESS)
    
End Function

' Liefert den Wert eines Eintrags, der durch Root, Schlüssel und Feld spezifiziert wird
Public Function GetValue(lRoot As Long, sKey As String, sField As String, vntValue As Variant) As Boolean
    
    Dim lResult&, keyhandle&, dwType&
    Dim zw&, puffergröße&, puffer$
    
    lResult = RegOpenKeyEx(lRoot, sKey, 0, KEY_READ, keyhandle)
    GetValue = (lResult = ERROR_SUCCESS)
    If lResult <> ERROR_SUCCESS Then Exit Function ' Schlüssel existiert nicht
    lResult = RegQueryValueEx(keyhandle, sField, 0&, dwType, ByVal 0&, puffergröße)
    GetValue = (lResult = ERROR_SUCCESS)
    If lResult <> ERROR_SUCCESS Then Exit Function ' Feld existiert nicht
    Select Case dwType
        Case REG_SZ       ' nullterminierter String
            puffer = Space$(puffergröße + 1)
            lResult = RegQueryValueEx(keyhandle, sField, 0&, dwType, ByVal puffer, puffergröße)
            GetValue = (lResult = ERROR_SUCCESS)
            If lResult <> ERROR_SUCCESS Then Exit Function ' Fehler beim auslesen des Feldes
            vntValue = ZTrim(puffer)
        Case REG_DWORD     ' 32-Bit Number   !!!! Word
            puffergröße = 4      ' = 32 Bit
            lResult = RegQueryValueEx(keyhandle, sField, 0&, dwType, zw, puffergröße)
            GetValue = (lResult = ERROR_SUCCESS)
            If lResult <> ERROR_SUCCESS Then Exit Function ' Fehler beim auslesen des Feldes
            vntValue = zw
        ' Hier könnten auch die weiteren Datentypen behandelt werden, soweit dies sinnvoll ist
    End Select
    If lResult = ERROR_SUCCESS Then RegCloseKey keyhandle
    GetValue = True
    
End Function

Private Function CreateKey(lRoot&, sNewkey$, sClass$) As Boolean

    'MD-Marker , Function wird nicht aufgerufen
    
'    Dim lResult&, keyhandle&
'    Dim Action&
'
'    lResult = RegCreateKeyEx(Root, Newkey, 0, Class, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, keyhandle, Action)
'    If lResult = ERROR_SUCCESS Then
'        If RegFlushKey(keyhandle) = ERROR_SUCCESS Then RegCloseKey keyhandle
'    Else
'         CreateKey = False
'        Exit Function
'    End If
'    CreateKey = (Action = REG_CREATED_NEW_KEY)
    
End Function

Private Function SetValue(lRoot As Long, sKey As String, sField As String, vntValue As Variant) As Boolean
    
    'MD-Marker , Function wird nicht aufgerufen
    
'    Dim lResult&, keyhandle&
'    Dim s$, l&
'
'    lResult = RegOpenKeyEx(lRoot, sKey, 0, KEY_ALL_ACCESS, keyhandle)
'    If lResult <> ERROR_SUCCESS Then
'        SetValue = False
'        Exit Function
'    End If
'    Select Case VarType(vntValue)
'        Case vbInteger, vbLong
'            l = CLng(vntValue)
'            lResult = RegSetValueEx_DWord(keyhandle, sField, 0, REG_DWORD, l, 4)
'        Case vbString
'            s = CStr(vntValue)
'            lResult = RegSetValueEx_String(keyhandle, sField, 0, REG_SZ, s, Len(s) + 1)    ' +1 für die Null am Ende
'
'        ' Hier können noch weitere Datentypen umgewandelt bzw. gespeichert werden
'    End Select
'    RegCloseKey keyhandle
'    SetValue = (lResult = ERROR_SUCCESS)
End Function

Private Function DeleteKey(lRoot&, sKey$) As Boolean

    'MD-Marker , Function wird nicht aufgerufen
    'Dim lResult&
    
    'lResult = RegDeleteKey(Root, Key)
    'DeleteKey = (lResult = ERROR_SUCCESS)
End Function

Private Function DeleteValue(lRoot&, sKey$, sField$) As Boolean

    'MD-Marker , Function wird nicht aufgerufen
    'Dim lResult&, keyhandle&
    
    'lResult = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, keyhandle)
    'If lResult <> ERROR_SUCCESS Then
    '    DeleteValue = False
    '    Exit Function
    'End If
    'lResult = RegDeleteValue(keyhandle, Field)
    'DeleteValue = (lResult = ERROR_SUCCESS)
    'RegCloseKey keyhandle
End Function
Private Function ZTrim(vZString As Variant) As Variant
  Dim nullpos&
  ZTrim = ""
  If IsNull(vZString) Then Exit Function
  nullpos& = InStr(vZString, Chr$(0))
  If nullpos = 0 Then
    ZTrim = vZString
    Exit Function
  End If
  If nullpos = 1 Then Exit Function
  ZTrim = Trim$(Left$(vZString, nullpos - 1))
End Function

