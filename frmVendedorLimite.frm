VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendedorLimite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Límite para vendedores"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendedorLimite.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12345
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Asignar limite a vendedor"
      TabPicture(0)   =   "frmVendedorLimite.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblIdVendedor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGrabar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtVendedor"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancelar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdServer"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Listado de vendedores"
      TabPicture(1)   =   "frmVendedorLimite.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "lvSearch"
      Tab(1).Control(2)=   "txtSearch"
      Tab(1).Control(3)=   "cmdSearch"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdServer 
         Caption         =   "Server"
         Enabled         =   0   'False
         Height          =   600
         Left            =   8160
         Picture         =   "frmVendedorLimite.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   600
         Left            =   10800
         Picture         =   "frmVendedorLimite.frx":146C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6840
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   360
         Left            =   -63840
         Picture         =   "frmVendedorLimite.frx":1BD6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   442
         Width           =   630
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   -73680
         TabIndex        =   0
         Top             =   480
         Width           =   9855
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   11
         Top             =   840
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   11668
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   1320
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   760
         Visible         =   0   'False
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2655
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1750
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame1 
         Height          =   5895
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   11895
         Begin VB.CommandButton cmdDel 
            Enabled         =   0   'False
            Height          =   360
            Left            =   11280
            Picture         =   "frmVendedorLimite.frx":1F60
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   360
            Left            =   10800
            Picture         =   "frmVendedorLimite.frx":22EA
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   600
            Width           =   495
         End
         Begin MSComctlLib.ListView lvData 
            Height          =   4815
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   8493
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.TextBox txtLimite 
            Height          =   315
            Left            =   6600
            TabIndex        =   6
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtArticulo 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label lblIdArticulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label5"
            Height          =   195
            Left            =   2160
            TabIndex        =   20
            Tag             =   "X"
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   6600
            TabIndex        =   19
            Top             =   360
            Width           =   510
         End
         Begin VB.Label lblPrecio 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4800
            TabIndex        =   18
            Top             =   600
            Width           =   1755
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Precio"
            Height          =   195
            Left            =   4800
            TabIndex        =   17
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Artículo"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   645
         End
      End
      Begin VB.TextBox txtVendedor 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   600
         Left            =   9480
         Picture         =   "frmVendedorLimite.frx":2674
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   21
         Top             =   525
         Width           =   900
      End
      Begin VB.Label lblIdVendedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   195
         Left            =   4440
         TabIndex        =   14
         Tag             =   "X"
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmVendedorLimite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vBuscarC As Boolean 'variable para la busqueda de vendedores
Private vBuscarA As Boolean 'variable para la busqueda de articulos
Private loc_key  As Integer
Private loc_keya  As Integer

Private Sub cmdAdd_Click()

    If Len(Trim(Me.lblIdArticulo.Caption)) = 0 Then
        MsgBox "Debe elegir un Articulo", vbCritical, Pub_Titulo
        Me.txtArticulo.SetFocus
        Exit Sub

    End If

    If Len(Trim(Me.txtLimite.Text)) = 0 Then
        MsgBox "Debe ingresar el límite", vbCritical, Pub_Titulo
        Me.txtLimite.SetFocus
        Exit Sub

    End If

    If Not IsNumeric(Me.txtLimite.Text) Then
        MsgBox "Límite ingresado incorrecto, debe ser número", vbCritical, Pub_Titulo
        Me.txtLimite.SetFocus
        Exit Sub

    End If
    
    If Val(Me.txtLimite.Text) <= 0 Then
        MsgBox "Limite ingresado no puede ser menor o igual a cero", vbCritical, Pub_Titulo
        Me.txtLimite.SetFocus
        Exit Sub
    End If

    Dim itemx As Object

    If Me.lvData.ListItems.count = 0 Then
 
        Set itemx = Me.lvData.ListItems.Add(, , Me.lblIdArticulo.Caption)
        itemx.SubItems(1) = Trim(Me.txtArticulo.Text)
        itemx.SubItems(2) = Me.txtLimite.Text
    Else

        Dim vEncontrado As Boolean

        vEncontrado = False

        For Each itemx In Me.lvData.ListItems

            If itemx.Text = Me.lblIdArticulo.Caption Then
                vEncontrado = True
                Exit For

            End If

        Next

        If vEncontrado = False Then
            Set itemx = Me.lvData.ListItems.Add(, , Me.lblIdArticulo.Caption)
            itemx.SubItems(1) = Trim(Me.txtArticulo.Text)
            itemx.SubItems(2) = Me.txtLimite.Text
        Else
            MsgBox "Producto ya agregado", vbCritical, Pub_Titulo

        End If

    End If

    Me.lblIdArticulo.Caption = ""
    Me.txtArticulo.Text = ""
    Me.txtLimite.Text = ""
    Me.lblPrecio.Caption = ""
    Me.txtArticulo.SetFocus

End Sub

Private Sub cmdCancelar_Click()
LimpiarControles Me
Me.lvData.ListItems.Clear
Me.ListView1.Visible = False
Me.ListView2.Visible = False
Me.ListView1.ListItems.Clear
Me.ListView2.ListItems.Clear
Me.SSTab1.tab = 1
Me.txtSearch.SetFocus
End Sub

Private Sub cmdDel_Click()
Me.lvData.ListItems.Remove Me.lvData.SelectedItem.Index
Me.cmdDel.Enabled = False
End Sub

Private Sub cmdGrabar_Click()

    If Len(Trim(Me.lblIdVendedor.Caption)) = 0 Then
        MsgBox "Debe ingresar el vendedor", vbCritical, Pub_Titulo
        Exit Sub

    End If
    
    'If Me.lvData.ListItems.count = 0 Then
      '  MsgBox "Debe agregar articulos.", vbCritical, Pub_Titulo
       ' Me.txtArticulo.SetFocus
       ' Exit Sub
    'End If

    On Error GoTo almacena

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_VENDEDOR_LIMITE_REGISTER]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@idvendedor", adInteger, adParamInput, , Me.lblIdVendedor.Caption)
    
    Dim strItems As String

    Dim f        As Integer
    
    If Me.lvData.ListItems.count <> 0 Then
        strItems = "<r>"

        For f = 1 To Me.lvData.ListItems.count
            strItems = strItems & "<d "
            strItems = strItems & "ida=""" & Me.lvData.ListItems(f).Text & """ "
            strItems = strItems & "lim=""" & Me.lvData.ListItems(f).SubItems(2) & """ "
            strItems = strItems & "/>"
        Next
        strItems = strItems & "</r>"

    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@items", adVarChar, adParamInput, 4000, strItems)

    Dim oRSdatos As ADODB.Recordset

    Set oRSdatos = oCmdEjec.Execute
  
    If Not oRSdatos.EOF Then
        If oRSdatos!Codigo = 0 Then
           
            cmdServer_Click
            cmdCancelar_Click
             LimpiarControles Me
            Me.ListView1.ListItems.Clear
            Me.ListView2.ListItems.Clear
            Me.lvData.ListItems.Clear
            'MsgBox oRSdatos!Mensaje, vbInformation, Pub_Titulo
        Else
            MsgBox oRSdatos!mensaje, vbCritical, Pub_Titulo

        End If

    End If
    
    Exit Sub
almacena:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub cmdSearch_Click()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_VENDEDOR_SEARCH]"
Dim oRSdatos As ADODB.Recordset

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adVarChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 40, Me.txtSearch.Text)

Set oRSdatos = oCmdEjec.Execute

Me.lvSearch.ListItems.Clear
Dim itemx As Object
Do While Not oRSdatos.EOF
    Set itemx = Me.lvSearch.ListItems.Add(, , oRSdatos!cod)
    itemx.SubItems(1) = Trim(oRSdatos!nom)
    oRSdatos.MoveNext
Loop
End Sub

Private Sub cmdServer_Click()

    If Len(Trim(Me.lblIdVendedor.Caption)) = 0 Then
        MsgBox "Debe ingresar el vendedor", vbCritical, Pub_Titulo
        Exit Sub

    End If
    
'    If Me.lvData.ListItems.count = 0 Then
'        MsgBox "Debe agregar articulos.", vbCritical, Pub_Titulo
'        Me.txtArticulo.SetFocus
'        Exit Sub
'
'    End If

    MousePointer = vbHourglass
'DATOS CLOUD
Dim c_Server As String, c_DataBase As String, c_User As String, c_Pass As String

c_Server = Leer_Ini(App.Path & "\config.ini", "C_SERVER", "c:\")
c_DataBase = Leer_Ini(App.Path & "\config.ini", "C_DATABASE", "c:\")
c_User = Leer_Ini(App.Path & "\config.ini", "C_USER", "c:\")
c_Pass = Leer_Ini(App.Path & "\config.ini", "C_PASS", "c:\")
'FIN CLOUD

    On Error GoTo server

    Dim oCnnRemoto As New ADODB.Connection

    oCnnRemoto.CursorLocation = adUseClient
    oCnnRemoto.Provider = "SQLOLEDB.1"
    oCnnRemoto.Open "Server=" + c_Server + ";Database=" + c_DataBase + ";Uid=" + c_User + ";Pwd=" + c_Pass + ";"

    Dim oCmdRemoto As New ADODB.Command

    oCmdRemoto.ActiveConnection = oCnnRemoto
    oCmdRemoto.CommandText = "[dbo].[USP_VENDEDOR_LIMITE_REGISTER]"
    oCmdRemoto.CommandType = adCmdStoredProc

    
    oCmdRemoto.Parameters.Append oCmdEjec.CreateParameter("@idvendedor", adInteger, adParamInput, , Me.lblIdVendedor.Caption)
    
    Dim strItems As String

    Dim f        As Integer
    
    If Me.lvData.ListItems.count <> 0 Then
        strItems = "<r>"

        For f = 1 To Me.lvData.ListItems.count
            strItems = strItems & "<d "
            strItems = strItems & "ida=""" & Me.lvData.ListItems(f).Text & """ "
            strItems = strItems & "lim=""" & Me.lvData.ListItems(f).SubItems(2) & """ "
            strItems = strItems & "/>"
        Next
        strItems = strItems & "</r>"

    End If

    oCmdRemoto.Parameters.Append oCmdEjec.CreateParameter("@items", adVarChar, adParamInput, 4000, strItems)

    Dim oRSdatos As ADODB.Recordset

    Set oRSdatos = oCmdRemoto.Execute
  
    If Not oRSdatos.EOF Then
        If oRSdatos!Codigo = 0 Then
            LimpiarControles Me
            Me.ListView1.ListItems.Clear
            Me.ListView2.ListItems.Clear
            Me.lvData.ListItems.Clear
            cmdCancelar_Click
            Me.cmdServer.Enabled = False
            MsgBox oRSdatos!mensaje, vbInformation, Pub_Titulo
        Else
            MsgBox oRSdatos!mensaje, vbCritical, Pub_Titulo

        End If

    End If

    oCnnRemoto.Close
    MousePointer = vbDefault
    Exit Sub
server:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.SSTab1.tab = 1
vBuscarC = False
CentrarFormulario MDIForm1, Me
ConfiguraLV

End Sub



Private Sub lvData_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdDel.Enabled = True
End Sub

Private Sub lvSearch_DblClick()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_VENDEDOR_LIMITE_FILL]"

    Dim itemx     As Object

    Dim orsResult As ADODB.Recordset

    Set orsResult = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lvSearch.SelectedItem.Text))

    If orsResult.EOF Then
        MsgBox "Vendedor no asignado", vbCritical, Pub_Titulo
        Exit Sub

    End If

    Me.lvData.ListItems.Clear

    Do While Not orsResult.EOF
        Me.txtVendedor.Text = orsResult!nomvend
        Me.lblIdVendedor.Caption = orsResult!idevend
        
        Set itemx = Me.lvData.ListItems.Add(, , orsResult!ideprod)
        itemx.SubItems(1) = orsResult!nomprod
        itemx.SubItems(2) = orsResult!limite
        orsResult.MoveNext
    Loop
    cmdServer.Enabled = True
    Me.SSTab1.tab = 0

End Sub

Private Sub txtArticulo_Change()
vBuscarA = True
Me.lblIdArticulo.Caption = ""
End Sub

Private Sub txtArticulo_GotFocus()
vBuscarA = True
Me.ListView1.Visible = False
End Sub

Private Sub txtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_keya = loc_keya + 1

        If loc_keya > ListView2.ListItems.count Then loc_keya = ListView2.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 38 Then
        loc_keya = loc_keya - 1

        If loc_keya < 1 Then loc_keya = 1
        GoTo POSICION
    End If

    If KeyCode = 34 Then
        loc_keya = loc_keya + 17

        If loc_keya > ListView2.ListItems.count Then loc_keya = ListView2.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 33 Then
        loc_keya = loc_keya - 17

        If loc_keya < 1 Then loc_keya = 1
        GoTo POSICION
    End If

    If KeyCode = 27 Then
        Me.ListView2.Visible = False
        Me.txtArticulo.Text = ""
        Me.lblIdArticulo.Caption = ""
    End If

    GoTo fin
POSICION:
    ListView2.ListItems.Item(loc_keya).Selected = True
    ListView2.ListItems.Item(loc_keya).EnsureVisible
    'txtRS.Text = Trim(ListView2.ListItems.Item(loc_keya).Text) & " "
    txtArticulo.SelStart = Len(txtArticulo.Text)
    
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscarA Then
            Me.ListView2.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[dbo].[USP_PRODUCTO_SEARCH]"
            Set oRsPago = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtArticulo.Text))

            Dim Item As Object
        
            If Not oRsPago.EOF Then

                Do While Not oRsPago.EOF
                    Set Item = Me.ListView2.ListItems.Add(, , oRsPago!cod)
                    Item.SubItems(1) = Trim(oRsPago!prod)
                    Item.SubItems(2) = Trim(oRsPago!PRE)
                    oRsPago.MoveNext
                Loop

                Me.ListView2.Visible = True
                Me.ListView2.ListItems(1).Selected = True
                loc_keya = 1
                Me.ListView2.ListItems(1).EnsureVisible
                vBuscarA = False
            
            End If
        
        Else
            
            Me.txtArticulo.Text = Me.ListView2.ListItems(loc_keya).SubItems(1)
            Me.lblIdArticulo.Caption = Me.ListView2.ListItems(loc_keya).Text
            Me.lblPrecio.Caption = Me.ListView2.ListItems(loc_keya).SubItems(2)
            Me.ListView2.Visible = False
            Me.txtLimite.SetFocus
'            Me.lvDetalle.SetFocus
        End If
    End If
End Sub

Private Sub txtlimite_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAdd_Click
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdSearch_Click
End Sub

Private Sub txtVendedor_Change()
vBuscarC = True
Me.lblIdVendedor.Caption = ""
End Sub

Private Sub txtVendedor_GotFocus()
vBuscarC = True
Me.ListView2.Visible = False
End Sub

Private Sub txtVendedor_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 27 Then
        Me.ListView1.Visible = False
        Me.txtVendedor.Text = ""
        Me.lblIdVendedor.Caption = ""
    End If

    GoTo fin
POSICION:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtVendedor.SelStart = Len(txtVendedor.Text)
    
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscarC Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[dbo].[USP_VENDEDOR_SEARCH]"
            Set oRsPago = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtVendedor.Text))

            Dim Item As Object
        
            If Not oRsPago.EOF Then

                Do While Not oRsPago.EOF
                    Set Item = Me.ListView1.ListItems.Add(, , oRsPago!cod)
                    Item.SubItems(1) = Trim(oRsPago!nom)
                    oRsPago.MoveNext
                Loop

                Me.ListView1.Visible = True
                Me.ListView1.ListItems(1).Selected = True
                loc_key = 1
                Me.ListView1.ListItems(1).EnsureVisible
                vBuscarC = False
            
            End If
        
        Else
            
            Me.txtVendedor.Text = Me.ListView1.ListItems(loc_key).SubItems(1)
            Me.lblIdVendedor.Caption = Me.ListView1.ListItems(loc_key).Text
            Me.ListView1.Visible = False
            Me.txtArticulo.SetFocus
'            Me.lvDetalle.SetFocus
        End If
    End If

End Sub

Private Sub ConfiguraLV()
With Me.ListView1
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .MultiSelect = False
End With
With Me.ListView2
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .ColumnHeaders.Add , , "precio", 0
    .MultiSelect = False
End With
With Me.lvData
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Articulo", 5000
    .ColumnHeaders.Add , , "Limite", 1500
    .MultiSelect = False
End With
With Me.lvSearch
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Nombres", 5000
    .MultiSelect = False
End With
End Sub
