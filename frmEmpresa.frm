VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Empresas"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8805
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   1164
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Nuevo"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Guardar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Editar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Cancelar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Eliminar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmEmpresa.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lvListado"
      Tab(0).Control(1)=   "txtSearch"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Empresa"
      TabPicture(1)   =   "frmEmpresa.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtDenominacion"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cboPrincipal"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtIdEmpresa"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtIdEmpresa 
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox cboPrincipal 
         Height          =   315
         ItemData        =   "frmEmpresa.frx":0038
         Left            =   2760
         List            =   "frmEmpresa.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txtDenominacion 
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   2760
         Width           =   4215
      End
      Begin MSComctlLib.ListView lvListado 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   2
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8070
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
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   8295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Es Principal:"
         Height          =   195
         Left            =   1545
         TabIndex        =   9
         Top             =   3420
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
         Height          =   195
         Left            =   1305
         TabIndex        =   7
         Top             =   2850
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Identificador:"
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   2250
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
Estado_Botones InicializarFormulario
ConfigurarLV
MostrarEmpresas
End Sub

Private Sub lvListado_DblClick()
EnviarDatosEdicion
End Sub
Private Sub EnviarDatosEdicion()
    Me.txtIdEmpresa.Text = Me.lvListado.SelectedItem.Text
    Me.txtDenominacion.Text = Me.lvListado.SelectedItem.SubItems(1)
    If Me.lvListado.SelectedItem.SubItems(2) = "SI" Then
        Me.cboPrincipal.ListIndex = 1
    Else
     Me.cboPrincipal.ListIndex = 0
    End If
    
End Sub

Private Sub lvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Estado_Botones AntesDeActualizar

    If Me.lvListado.SelectedItem.SubItems(2) = "SI" Then
        Me.Toolbar1.Buttons(5).Caption = "&Desactiva"
    Else
        Me.Toolbar1.Buttons(5).Caption = "&Activa"

    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            VNuevo = True
           
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            Me.cboPrincipal.ListIndex = 1
        Case 2 'GRABAR
            grabarEmpresa

        Case 3 'MODIFICAR
            VNuevo = False
            ActivarControles Me
            Estado_Botones Editar
            EnviarDatosEdicion
            Me.txtIdEmpresa.Enabled = False
            Me.txtDenominacion.SetFocus

        Case 4 'CANCELAR
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvListado.Enabled = True
            Me.txtSearch.Enabled = True

        Case 5 'ACTIVAR/DESACTIVAR

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, gNombreProyecto) = vbYes Then

                On Error GoTo ESTADO

                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "[dbo].[USP_EMPRESA_ESTADO]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , Me.lvListado.SelectedItem.Text)

                If Me.lvListado.SelectedItem.SubItems(2) = "SI" Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , False)
                Else
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , True)

                End If

                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)

                oCmdEjec.Execute
                MostrarEmpresas
                LimpiarControles Me
                Estado_Botones Eliminar
               
                Exit Sub

ESTADO:
                MsgBox Err.Description, vbInformation, NombreProyecto

            End If

        Case 6 'eliminar

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, gNombreProyecto) = vbYes Then

                On Error GoTo Elimina

                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "[dbo].[USP_EMPRESA_ELIMINAR]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , Me.lvListado.SelectedItem.Text)
    

                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)

                oCmdEjec.Execute
                MostrarEmpresas
                LimpiarControles Me
                Estado_Botones Eliminar
               
                Exit Sub

Elimina:
                MsgBox Err.Description, vbInformation, NombreProyecto

            End If

    End Select

End Sub

Private Sub Estado_Botones(val As Valores)

    Select Case val

        Case InicializarFormulario, grabar, cancelar, Eliminar
            Me.Toolbar1.Buttons(1).Enabled = True
            Me.Toolbar1.Buttons(2).Enabled = False
            Me.Toolbar1.Buttons(3).Enabled = False
            Me.Toolbar1.Buttons(4).Enabled = False
            Me.Toolbar1.Buttons(5).Enabled = False
            Me.Toolbar1.Buttons(6).Enabled = False
            Me.SSTab1.tab = 0
        
        Case Nuevo, Editar
            Me.Toolbar1.Buttons(1).Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = True
            Me.Toolbar1.Buttons(3).Enabled = False
            Me.Toolbar1.Buttons(4).Enabled = True
            Me.Toolbar1.Buttons(5).Enabled = False
Me.Toolbar1.Buttons(6).Enabled = False
            Me.SSTab1.tab = 1
            Me.txtIdEmpresa.SetFocus

        Case buscar
            Me.Toolbar1.Buttons(1).Enabled = True
            Me.Toolbar1.Buttons(2).Enabled = False
            Me.Toolbar1.Buttons(3).Enabled = False
            Me.Toolbar1.Buttons(4).Enabled = False

            Me.SSTab1.tab = 1

        Case AntesDeActualizar
            Me.Toolbar1.Buttons(1).Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = False
            Me.Toolbar1.Buttons(3).Enabled = True
            Me.Toolbar1.Buttons(4).Enabled = True
             Me.Toolbar1.Buttons(5).Enabled = True
            Me.Toolbar1.Buttons(6).Enabled = True
            Me.SSTab1.tab = 0

    End Select

End Sub

Sub grabarEmpresa()

    If Len(Trim(Me.txtIdEmpresa.Text)) = 0 Then
        MsgBox "Debe ingregar el Identificador de Empresa.", vbCritical, Pub_Titulo
        Me.txtIdEmpresa.SetFocus
    
        
        Exit Sub

    End If
    
    If Not IsNumeric(Me.txtIdEmpresa.Text) Then
        MsgBox "Identificador ingresado es incorrecto.", vbCritical, Pub_Titulo
            Me.txtIdEmpresa.SelStart = 0
        Me.txtIdEmpresa.SelLength = Len(Me.txtIdEmpresa.Text)
        Me.txtIdEmpresa.SetFocus
        Exit Sub
    End If

    If Len(Trim(Me.txtDenominacion.Text)) = 0 Then
        MsgBox "Debe ingresar la denominación de la Empresa", vbCritical, Pub_Titulo
        Me.txtDenominacion.SetFocus
        Exit Sub

    End If

    On Error GoTo xGraba:

    LimpiaParametros oCmdEjec

    Dim orsGraba As ADODB.Recordset

    If VNuevo Then
        oCmdEjec.CommandText = "[dbo].[USP_EMPRESA_REGISTRAR]"
    Else
        oCmdEjec.CommandText = "[dbo].[USP_EMPRESA_ACTUALIZAR]"
        

    End If
   oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , Me.txtIdEmpresa.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DENOMINACION", adVarChar, adParamInput, 200, Trim(Me.txtDenominacion.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DEFECTO", adBoolean, adParamInput, , Me.cboPrincipal.ListIndex)
    Set orsGraba = oCmdEjec.Execute
   
    If Not orsGraba.EOF Then
        If orsGraba!Code = 0 Then
            MostrarEmpresas
            DesactivarControles Me
            Estado_Botones grabar
        Else
            MsgBox orsGraba!Message, vbCritical, Pub_Titulo

        End If

    End If
    
    Exit Sub
xGraba:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub ConfigurarLV()
With Me.lvListado
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Identificador", 1500
    .ColumnHeaders.Add , , "Denominación", 3500
    .ColumnHeaders.Add , , "Principal", 700
    .ColumnHeaders.Add , , "Activo", 700

End With
End Sub

Private Sub MostrarEmpresas()
    Me.lvListado.ListItems.Clear

    Dim datos As Object, dError As Boolean

    On Error GoTo tDatos

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_EMPRESA_LIST]"

    If Len(Trim(Me.txtSearch.Text)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 100, Me.txtSearch.Text)

    Dim orsDatos As New ADODB.Recordset

    Set orsDatos = oCmdEjec.Execute

    Dim itemx As Object

    Do While Not orsDatos.EOF
        Set itemx = Me.lvListado.ListItems.Add(, , orsDatos!idempresa)
        itemx.SubItems(1) = orsDatos!DENOMINACION

        If orsDatos!defecto Then
            itemx.SubItems(2) = "SI"
        Else
            itemx.SubItems(2) = "NO"

        End If

        If orsDatos!activo Then
            itemx.SubItems(3) = "SI"
        Else
            itemx.SubItems(3) = "NO"

        End If
    
        orsDatos.MoveNext
    Loop

    Exit Sub

tDatos:
    MsgBox Err.Description, vbCritical, gNombreProyecto

End Sub

Private Sub txtDenominacion_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
End Sub

Private Sub txtIdEmpresa_Change()
ValidarSoloNumeros Me.txtIdEmpresa
End Sub

Private Sub txtIdEmpresa_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then MostrarEmpresas
End Sub
