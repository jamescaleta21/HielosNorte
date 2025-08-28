Attribute VB_Name = "Module4"
' Añade esto al módulo si usas la API para deshabilitar el ListView
Private Declare Function EnableWindow Lib "User32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

'Public VReporte As New CRAXDRT.Report
Public Enum Valores
    InicializarFormulario
    Nuevo
    grabar
    cancelar
    Editar
    buscar
    AntesDeActualizar
    Eliminar           'LINEA NUEVA
    Desactivar
    Activar
End Enum

Public Declare Function DrawMenuBar Lib "User32" _
      (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" _
      (ByVal hMenu As Long) As Long
Public Declare Function GetSystemMenu Lib "User32" _
        (ByVal hwnd As Long, _
        ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "User32" _
        (ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
        
Public oCmdEjec As New ADODB.Command
Public objAES As New clsSecurity
Public objCryp As New clsCrypto

Public Sub CerrarConexion(esRemoto As Boolean)

    If esRemoto Then
        If oCnnRemoto.State = adStateOpen Then
            oCnnRemoto.Close
            Set oCnnRemoto = Nothing

        End If

    Else

        If Pub_ConnAdo.State = adStateOpen Then
            Pub_ConnAdo.Close
            'Set Pub_ConnAdo = Nothing

        End If

    End If

End Sub

Public Sub LimpiaParametros(oCmd As ADODB.Command, Optional esRemoto As Boolean = False)

    If esRemoto Then

        'DATOS CLOUD
        Dim strCadenaConexion As String
        
        Dim sSemilla As String
        sSemilla = LeerValorRegistro("Semilla")
        
        sSemilla = objCryp.DecryptText(sSemilla)
        
        sSemilla = objAES.Semilla(sSemilla)
        
        strCadenaConexion = LeerValorRegistro("ConexionCloud")
        strCadenaConexion = objAES.DeCodificarB64(strCadenaConexion, sSemilla)

'        c_Server = Leer_Ini(App.Path & "\config.ini", "C_SERVER", "c:\")
'        c_Server = objAES.DeCodificarB64(c_Server, sSemilla)
'
'        c_DataBase = Leer_Ini(App.Path & "\config.ini", "C_DATABASE", "c:\")
'        c_DataBase = objAES.DeCodificarB64(c_DataBase, sSemilla)
'
'        c_User = Leer_Ini(App.Path & "\config.ini", "C_USER", "c:\")
'        c_User = objAES.DeCodificarB64(c_User, sSemilla)
'
'        c_Pass = Leer_Ini(App.Path & "\config.ini", "C_PASS", "c:\")
'        c_Pass = objAES.DeCodificarB64(c_Pass, sSemilla)
        'FIN CLOUD
        If oCnnRemoto.State = adStateOpen Then oCnnRemoto.Close
        oCnnRemoto.Provider = "SQLOLEDB.1"
        'oCnnRemoto.Open "Server=" + c_Server + ";Database=" + c_DataBase + ";Uid=" + c_User + ";Pwd=" + c_Pass + ";"
        oCnnRemoto.Open strCadenaConexion
        
        oCmd.ActiveConnection = oCnnRemoto
        oCnnRemoto.CursorLocation = adUseClient
    Else

        Dim xCadena As String
        
        If Pub_ConnAdo.State = adStateClosed Then
        Pub_ConnAdo.Open
'            xCadena = "dsn=" & wdsn & ";uid=sa;pwd=" & cClave
'            Pub_ConnAdo.Provider = "SQLOLEDB.1"
'            Pub_ConnAdo.Open xCadena

        End If

        oCmd.ActiveConnection = Pub_ConnAdo
        Pub_ConnAdo.CursorLocation = adUseClient

    End If

    oCmd.CommandType = adCmdStoredProc

    For I = oCmd.Parameters.count - 1 To 0 Step -1
        oCmd.Parameters.Delete I
    Next

End Sub

Public Sub InhabilitarCerrar(ofrm As Form)
Dim hMenu As Long
Dim menuItemCount As Long
'Obtenemos un handle al menú de sistema del formulario
hMenu = GetSystemMenu(ofrm.hwnd, 0)
If hMenu Then
    'Obtenemos el número de elementos del menú
    menuItemCount = GetMenuItemCount(hMenu)
    'Eliminamos el elemento Cerrar, que es el último
    'Los elemento empiezan a numerarse en cero por lo que el
    'último es menuItemCount - 1
     Call RemoveMenu(hMenu, menuItemCount - 1, _
                      MF_REMOVE Or MF_BYPOSITION)
    'Eliminamos la barra de separación que hay justo antes de la opción Cerrar
    Call RemoveMenu(hMenu, menuItemCount - 2, _
                      MF_REMOVE Or MF_BYPOSITION)
    'Forzamos el redibujado del menú. Esto refresca la barra de título
    'y deja la X deshabilitada
    Call DrawMenuBar(ofrm.hwnd)
End If
End Sub


Public Sub MostrarErrores(xError As ErrObject)
MsgBox "Descripcion del Error: " & xError.Description & vbCrLf & _
"Origen del Error: " & xError.Source & vbCrLf & "Número de Error: " & xError.Number, vbCritical, NombreProyecto
End Sub


Public Sub LimpiarControles(Frm As Form)
   Dim I
   For I = 0 To Frm.Controls.count - 1
      If TypeOf Frm.Controls(I) Is TextBox Then
         Frm.Controls(I).Text = ""
      ElseIf TypeOf Frm.Controls(I) Is label And Frm.Controls(I).Tag = "X" Then
          Frm.Controls(I).Caption = ""
      ElseIf TypeOf Frm.Controls(I) Is ComboBox Then
Frm.Controls(I).ListIndex = -1
      End If
   Next I
End Sub

Public Sub ActivarControles(Frm As Form)
Dim j As Integer

For j = 0 To Frm.Controls.count - 1

    If TypeOf Frm.Controls(j) Is TextBox Then
        Frm.Controls(j).Enabled = True
    End If

    If TypeOf Frm.Controls(j) Is DataCombo Then
        Frm.Controls(j).Enabled = True
    End If

    If TypeOf Frm.Controls(j) Is ComboBox Then
        Frm.Controls(j).Enabled = True
    End If

    If TypeOf Frm.Controls(j) Is DTPicker Then
        Frm.Controls(j).Enabled = True
    End If

    If TypeOf Frm.Controls(j) Is MaskEdBox Then
        Frm.Controls(j).Enabled = True
    End If

    If TypeOf Frm.Controls(j) Is UpDown Or TypeOf Frm.Controls(j) Is OptionButton Then
        Frm.Controls(j).Enabled = True
    End If

   If TypeOf Frm.Controls(j) Is ListView And Frm.Controls(j).Tag = "X" Then
        Frm.Controls(j).Enabled = True
    End If
     If TypeOf Frm.Controls(j) Is CommandButton And Frm.Controls(j).Tag = "X" Then
        Frm.Controls(j).Enabled = True
    End If
Next
End Sub

Public Function Mayusculas(Caracter As Integer) As Integer
    'Para escribir en Mayusculas
    
    Mayusculas = Asc(UCase(Chr(Caracter)))
End Function

Public Sub DesactivarControles(Frm As Form)
    Dim j As Integer
    Dim ctrl As Control
    
    For j = 0 To Frm.Controls.count - 1
        Set ctrl = Frm.Controls(j)
        
        ' TextBox con Tag = "X"
        If TypeOf ctrl Is TextBox And ctrl.Tag = "X" Then ctrl.Enabled = False
        
        ' DataCombo
        If TypeOf ctrl Is DataCombo Then ctrl.Enabled = False
        
        ' ComboBox
        If TypeOf ctrl Is ComboBox Then ctrl.Enabled = False

        ' DTPicker
        If TypeOf ctrl Is DTPicker Then ctrl.Enabled = False
        
        ' MaskEdBox
        If TypeOf ctrl Is MaskEdBox Then
            ctrl.Enabled = False
        End If
        
        ' OptionButton
        If TypeOf ctrl Is OptionButton Then
            ctrl.Enabled = False
        End If
        
        ' ListView - Solución especial
        If TypeOf ctrl Is ListView Then
            ' Alternativa 1: Deshabilitar a través de contenedor (si existe)
            On Error Resume Next ' Por si no tiene hWnd
            EnableWindow ctrl.hwnd, False
            On Error GoTo 0
            
            ' Alternativa 2: Hacerlo de solo lectura
            ctrl.Enabled = False ' Algunas versiones sí permiten esto
            ctrl.LabelEdit = lvwNone
        End If
        
        ' CommandButton con Tag = "X"
        If TypeOf ctrl Is CommandButton And ctrl.Tag = "X" Then
            ctrl.Enabled = False
        End If
    Next
End Sub



