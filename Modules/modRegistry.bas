Attribute VB_Name = "modRegistry"
Option Explicit

' ====== Declaraciones de API ======
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long

' ====== Constantes ======
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_SZ = 1

Private Const SUBKEY_PATH = "SOFTWARE\JMSOFTWARE"

' -------------------------------------------------------------
' Función: Lee un valor de HKLM\SOFTWARE\JMSOFTWARE
' Recibe el nombre del valor (ej: "Semilla") y devuelve su contenido
' Si no existe, devuelve ""
' -------------------------------------------------------------
Public Function LeerValorRegistro(ByVal nombreValor As String) As String
    Dim hKey As Long, lResult As Long
    Dim sBuffer As String, lSize As Long, lType As Long
    
    LeerValorRegistro = ""
    
    ' Abrir clave (solo lectura)
    lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SUBKEY_PATH, 0, KEY_QUERY_VALUE, hKey)
    If lResult <> 0 Then Exit Function   ' Clave no existe
    
    ' Consultar tamaño del valor
    lResult = RegQueryValueEx(hKey, nombreValor, 0, lType, ByVal 0&, lSize)
    If lResult = 0 And lType = REG_SZ Then
        sBuffer = String(lSize, 0)
        lResult = RegQueryValueEx(hKey, nombreValor, 0, lType, ByVal sBuffer, lSize)
        If lResult = 0 Then
            LeerValorRegistro = Left$(sBuffer, lSize)
        End If
    End If
    
    RegCloseKey hKey
End Function

' -------------------------------------------------------------
' Procedimiento: Graba o crea un valor en HKLM\SOFTWARE\JMSOFTWARE
' -------------------------------------------------------------
Public Sub GrabarValorRegistro(ByVal nombreValor As String, ByVal nuevoValor As String)
    Dim hKey As Long, lResult As Long, lDisp As Long
    
    ' Crear o abrir clave JMSOFTWARE
    lResult = RegCreateKeyEx(HKEY_LOCAL_MACHINE, SUBKEY_PATH, 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, lDisp)
    If lResult <> 0 Then
        MsgBox "Error al acceder al registro. Ejecuta como Administrador.", vbCritical
        Exit Sub
    End If
    
    ' Establecer valor
    lResult = RegSetValueEx(hKey, nombreValor, 0, REG_SZ, ByVal nuevoValor, Len(nuevoValor))
    If lResult = 0 Then
        MsgBox "Valor '" & nombreValor & "' grabado correctamente.", vbInformation
    Else
        MsgBox "Error al grabar el valor '" & nombreValor & "'.", vbCritical
    End If
    
    RegCloseKey hKey
End Sub




