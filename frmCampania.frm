VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCampania 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Campañas"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCampania.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9135
   Begin MSComctlLib.ImageList ilCampania 
      Left            =   8040
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCampania.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCampania.frx":1064
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCampania.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCampania.frx":1798
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCampania.frx":1B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCampania.frx":1ECC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab stabCampania 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmCampania.frx":2266
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtSearch"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lvListado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Campaña"
      TabPicture(1)   =   "frmCampania.frx":2282
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "lblIde"
      Tab(1).Control(6)=   "dtpFin"
      Tab(1).Control(7)=   "txtNombre"
      Tab(1).Control(8)=   "dtpIni"
      Tab(1).Control(9)=   "txtMonto"
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtMonto 
         Height          =   300
         Left            =   -71400
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "X"
         Top             =   3720
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   300
         Left            =   -71400
         TabIndex        =   2
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   84606977
         CurrentDate     =   45478
      End
      Begin VB.TextBox txtNombre 
         Height          =   300
         Left            =   -71400
         TabIndex        =   1
         Tag             =   "X"
         Top             =   2280
         Width           =   3495
      End
      Begin MSComctlLib.ListView lvListado 
         Height          =   4215
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7435
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
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   7575
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   300
         Left            =   -71400
         TabIndex        =   3
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   84606977
         CurrentDate     =   45478
      End
      Begin VB.Label lblIde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -71400
         TabIndex        =   14
         Top             =   1680
         Width           =   2115
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto de Campaña:"
         Height          =   195
         Left            =   -73290
         TabIndex        =   13
         Top             =   3780
         Width           =   1740
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Termino:"
         Height          =   195
         Left            =   -73155
         TabIndex        =   12
         Top             =   3300
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio:"
         Height          =   195
         Left            =   -72915
         TabIndex        =   11
         Top             =   2820
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Campaña:"
         Height          =   195
         Left            =   -73185
         TabIndex        =   10
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Campaña:"
         Height          =   195
         Left            =   -73080
         TabIndex        =   9
         Top             =   1740
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campaña:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   645
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1111
      ButtonWidth     =   1482
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ilCampania"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Estado"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCampania"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private vPUNTO As Boolean 'variable para controld epunto sin utilizar ocx
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
ConfigurarLV
DesactivarControles Me
Estado_Botones InicializarFormulario
RealizarBusqueda
CentrarFormulario MDIForm1, Me

Me.dtpIni.Value = LK_FECHA_DIA
Me.dtpFin.Value = DateAdd("m", 1, LK_FECHA_DIA)
End Sub


Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbMenu.Buttons(1).Enabled = True
        Me.tbMenu.Buttons(2).Enabled = False
        Me.tbMenu.Buttons(3).Enabled = False
        Me.tbMenu.Buttons(4).Enabled = False
        Me.tbMenu.Buttons(5).Enabled = False
        Me.stabCampania.tab = 0
    Case Nuevo, Editar
        Me.tbMenu.Buttons(1).Enabled = False
        Me.tbMenu.Buttons(2).Enabled = True
        Me.tbMenu.Buttons(3).Enabled = False
        Me.tbMenu.Buttons(4).Enabled = True
        Me.lvListado.Enabled = False
        Me.txtSearch.Enabled = False
        Me.stabCampania.tab = 1
        Me.tbMenu.Buttons(5).Enabled = False
    Case buscar
        Me.tbMenu.Buttons(1).Enabled = True
        Me.tbMenu.Buttons(2).Enabled = False
        Me.tbMenu.Buttons(3).Enabled = False
        Me.tbMenu.Buttons(4).Enabled = False
        Me.stabCampania.tab = 0
    Case AntesDeActualizar
        Me.tbMenu.Buttons(1).Enabled = False
        Me.tbMenu.Buttons(2).Enabled = False
        Me.tbMenu.Buttons(3).Enabled = True
        Me.tbMenu.Buttons(4).Enabled = True
         'Me.tbmenu.Buttons(5).Enabled = True
        Me.stabCampania.tab = 1
End Select
End Sub


Private Sub ConfigurarLV()
With Me.lvListado
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Nombre", 3000
    .ColumnHeaders.Add , , "Inicio", 1500
    .ColumnHeaders.Add , , "Termino", 1400
    .ColumnHeaders.Add , , "Activo", 800
    .ColumnHeaders.Add , , "Monto", 800
End With
End Sub

Private Sub lvListado_Click()
    Me.tbMenu.Buttons(5).Enabled = True
    
    If Me.lvListado.SelectedItem.SubItems(3) = "NO" Then
        Me.tbMenu.Buttons(5).Caption = "&Activar"
        Me.tbMenu.Buttons(5).Image = 5
    Else
        Me.tbMenu.Buttons(5).Caption = "&Desactivar"
Me.tbMenu.Buttons(5).Image = 6
    End If

End Sub

Private Sub lvListado_DblClick()
If Me.lvListado.ListItems.count <> 0 Then Mandar_Datos
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            VNuevo = True
            Me.txtNombre.SetFocus

        Case 2 'GUARDAR
            LimpiaParametros oCmdEjec

            Dim orsProcesa      As ADODB.Recordset

            Dim orsProcesaCloud As ADODB.Recordset

            Dim xMensaje        As String

            Dim parts()         As String

            If Len(Trim(Me.txtNombre.Text)) = 0 Then
                MsgBox "Debe ingresar el Nombre de la Campaña", vbCritical, Pub_Titulo
                Me.txtNombre.SetFocus
            ElseIf Len(Trim(Me.txtMonto.Text)) = 0 Then
                MsgBox "Debe ingresar el monto para la Campaña.", vbCritical, Pub_Titulo
                Me.txtMonto.SetFocus
            Else

                On Error GoTo grabar

                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMBRE", adVarChar, adParamInput, 100, Trim(Me.txtNombre.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@INI", adChar, adParamInput, 8, ConvertirFechaFormat_yyyyMMdd(Me.dtpIni.Value))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FIN", adChar, adParamInput, 8, ConvertirFechaFormat_yyyyMMdd(Me.dtpFin.Value))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONTO", adDouble, adParamInput, , Me.txtMonto.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)

                If VNuevo Then
                    oCmdEjec.CommandText = "[dbo].[USP_CAMPANIA_REGISTRAR]"
                Else
                    oCmdEjec.CommandText = "[dbo].[USP_CAMPANIA_MODIFICAR]"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCAMPANIA", adInteger, adParamInput, , Me.lblIde.Caption)

                End If

                Set orsProcesa = oCmdEjec.Execute
                
                If Not orsProcesa.EOF Then
                    xMensaje = orsProcesa!mensaje
                    
                    If InStr(xMensaje, "=") > 0 Then
                        parts = Split(xMensaje, "=")

                        If parts(0) = 0 Then
                            MousePointer = vbDefault
                            MsgBox parts(1), vbInformation, Pub_Titulo
                            DesactivarControles Me
                            Estado_Botones grabar
                            
                            'GRABANDO EN CLOUD - inicio
                            LimpiaParametros oCmdEjec
                            oCmdEjec.Prepared = True
                            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                            MousePointer = vbHourglass

                            If VNuevo Then
                                oCmdEjec.CommandText = "[dbo].[USP_CAMPANIA_REGISTRAR_CLOUD]"
                                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCAMPANIA", adInteger, adParamInput, , orsProcesa!IDe)
                            Else
                                oCmdEjec.CommandText = "[dbo].[USP_CAMPANIA_MODIFICAR_CLOUD]"
                                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCAMPANIA", adInteger, adParamInput, , Me.lblIde.Caption)

                            End If
                            
                            Set orsProcesaCloud = oCmdEjec.Execute
                            
                            If Not orsProcesaCloud.EOF Then
                                MousePointer = vbDefault
                                xMensaje = orsProcesaCloud!mensaje
                                
                                If InStr(xMensaje, "=") > 0 Then
                                    parts = Split(xMensaje, "=")

                                    If parts(0) <> 0 Then
                                        MsgBox parts(1), vbCritical, Pub_Titulo

                                    End If

                                End If

                            End If
                            
                            'GRABANDO EN CLOUD - fin
                            
                            Me.lvListado.Enabled = True
                            Me.txtSearch.Enabled = True
                            RealizarBusqueda Me.txtSearch.Text
                        Else
                            MousePointer = vbDefault
                            MsgBox parts(1), vbCritical, Pub_Titulo

                        End If

                    Else
                        MousePointer = vbDefault
                        MsgBox "No se pudo leer el codigo de retorno.", vbInformation, Pub_Titulo

                    End If
                    
                End If

                MousePointer = vbDefault
                Exit Sub

grabar:
                MousePointer = vbDefault
                MsgBox Err.Description, vbInformation, Pub_Titulo

            End If

        Case 3 'MODIFICAR
            VNuevo = False
            Estado_Botones Editar
            ActivarControles Me
            'Me.txtCodigo.Enabled = False
            Me.txtNombre.SetFocus

        Case 4 'CANCELAR
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvListado.Enabled = True
            Me.txtSearch.Enabled = True
            Me.lvListado.SelectedItem.Selected = False

            '
        Case 5 'ELIMINAR

            Dim xEstado As Boolean
          
            If MsgBox("¿Desea continuar con la operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
          
            On Error GoTo cEstado
            
            If Me.lvListado.SelectedItem.SubItems(3) = "NO" Then
                xEstado = True
            Else
                xEstado = False

            End If

            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[dbo].[USP_CAMPANIA_ESTADO]"
    
            oCmdEjec.Prepared = True
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCAMPANIA", adInteger, adParamInput, , Me.lvListado.SelectedItem.Tag)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , xEstado)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)
                
            Dim orsEstado As ADODB.Recordset

            Dim oRScloud  As ADODB.Recordset

            MousePointer = vbHourglass

            Set orsEstado = oCmdEjec.Execute

            Dim aParts() As String

            MousePointer = vbDefault

            If Not orsEstado.EOF Then
                xMensaje = orsEstado!mensaje
                    
                If InStr(xMensaje, "=") > 0 Then
                    aParts = Split(xMensaje, "=")

                    If aParts(0) = 0 Then
                        MsgBox aParts(1), vbInformation, Pub_Titulo
                        MousePointer = vbHourglass
                        'ENVIANDO A CLOUD - inicio
                        LimpiaParametros oCmdEjec
                        oCmdEjec.CommandText = "[dbo].[USP_CAMPANIA_MODIFICAR_CLOUD]"
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCAMPANIA", adInteger, adParamInput, , Me.lvListado.SelectedItem.Tag)
                                
                        Set oRScloud = oCmdEjec.Execute
                        MousePointer = vbDefault

                        If Not oRScloud.EOF Then
                            xMensaje = oRScloud!mensaje
                                    
                            If InStr(xMensaje, "=") > 0 Then
                                aParts = Split(xMensaje, "=")
                                        
                                If aParts(0) <> 0 Then
                                    MsgBox aParts(1), vbCritical, Pub_Titulo

                                End If

                            End If

                        End If

                        'ENVIANDO A CLOUD - fin
                        
                        DesactivarControles Me
                        Estado_Botones grabar
                                
                        Me.lvListado.Enabled = True
                        Me.txtSearch.Enabled = True
                        RealizarBusqueda Me.txtSearch.Text
                    Else
                        MousePointer = vbDefault
                        MsgBox aParts(1), vbCritical, Pub_Titulo

                    End If

                Else
                    MousePointer = vbDefault
                    MsgBox "No se pudo leer el codigo de retorno.", vbCritical, Pub_Titulo

                End If

            End If

            MousePointer = vbDefault
            Exit Sub
            
cEstado:
            MousePointer = vbDefault
            MsgBox Err.Description, vbCritical, Pub_Titulo
            
            '
            '            On Error GoTo elimina
            '
            '            If MsgBox("¿Desea continuar con la operación.?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            '                LimpiaParametros oCmdEjec
            '                oCmdEjec.CommandText = "SP_FAMILIA_ELIMINAR"
            '                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            '                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adBigInt, adParamInput, , Me.LVLISTADO.SelectedItem.Tag)
            '                oCmdEjec.Execute
            '
            '                Me.LVLISTADO.ListItems.Remove Me.LVLISTADO.SelectedItem.Index
            '                Me.tbFamilia.Buttons(5).Enabled = False
            '
            '                MsgBox "Datos Eliminados Correctamente", vbInformation, Pub_Titulo
            '
            '            End If
            '
            '            Exit Sub
            '
            'elimina:
            '            MsgBox Err.Description, vbCritical, Pub_Titulo
            
    End Select

End Sub

Private Sub RealizarBusqueda(Optional vSearch As String = "")
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_CAMPANIA_SEARCH]"
    Me.lvListado.ListItems.Clear

    Dim ORSf As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    If Len(Trim(Me.txtSearch.Text)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 200, Trim(Me.txtSearch.Text))
    End If
    
    Set ORSf = oCmdEjec.Execute
    
    Do While Not ORSf.EOF

        With Me.lvListado.ListItems.Add(, , Trim(ORSf!nom))
            .Tag = Trim(ORSf!IDe)
            .SubItems(1) = ORSf!ini
            .SubItems(2) = ORSf!fin
            .SubItems(3) = ORSf!ACTIVO
            .SubItems(4) = ORSf!MONTO
        End With
        ORSf.MoveNext
    Loop

End Sub

Sub Mandar_Datos()
Me.tbMenu.Buttons(5).Enabled = False
    With Me.lvListado
        Me.lblIde.Caption = .SelectedItem.Tag
        Me.txtNombre.Text = .SelectedItem.Text
        Me.dtpIni.Value = .SelectedItem.SubItems(1)
        Me.dtpFin.Value = .SelectedItem.SubItems(2)
        Me.txtMonto.Text = .SelectedItem.SubItems(4)
        
        Estado_Botones AntesDeActualizar
    End With

End Sub

Private Sub txtMonto_Change()
If InStr(Me.txtMonto.Text, ".") Then
    vPUNTO = True
Else
    vPUNTO = False
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If KeyAscii = 46 Then
        If vPUNTO Or Len(Trim(Me.txtMonto.Text)) = 0 Then
            KeyAscii = 0

        End If

    End If
    
  
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then RealizarBusqueda
End Sub
