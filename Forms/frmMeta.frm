VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmMeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metas"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
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
   ScaleHeight     =   9660
   ScaleWidth      =   12045
   Begin MSComctlLib.ImageList ilMeta 
      Left            =   12240
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":039A
            Key             =   "save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":0734
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":0ACE
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":0E68
            Key             =   "inactive"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":1402
            Key             =   "active"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":199C
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":1D36
            Key             =   "meta"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar mtbMeta 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   1164
      ButtonWidth     =   1879
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Guardar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Desactivar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Activar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Eliminar"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTMeta 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmMeta.frx":20D0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lvDatos"
      Tab(0).Control(1)=   "txtSearch"
      Tab(0).Control(2)=   "Label1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Meta"
      TabPicture(1)   =   "frmMeta.frx":20EC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FraCabecera"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame FraDetalle 
         Height          =   5775
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   11535
         Begin VB.CommandButton cmdDel 
            Height          =   360
            Left            =   9480
            Picture         =   "frmMeta.frx":2108
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1920
            Width           =   990
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   360
            Left            =   9480
            Picture         =   "frmMeta.frx":2492
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1440
            Width           =   990
         End
         Begin MSComctlLib.ListView lvDetalle 
            Height          =   4455
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   7858
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.TextBox txtImporte 
            Height          =   375
            Left            =   6720
            TabIndex        =   9
            Tag             =   "X"
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtCantidad 
            Height          =   375
            Left            =   1680
            TabIndex        =   8
            Tag             =   "X"
            Top             =   720
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DatProducto 
            Height          =   360
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
            Height          =   240
            Left            =   5760
            TabIndex        =   24
            Top             =   787
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
            Height          =   240
            Left            =   615
            TabIndex        =   23
            Top             =   787
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Producto:"
            Height          =   240
            Left            =   600
            TabIndex        =   22
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.Frame FraCabecera 
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   11535
         Begin VB.TextBox txtDescripcion 
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Tag             =   "X"
            Top             =   840
            Width           =   7335
         End
         Begin MSMask.MaskEdBox mebFecIni 
            Height          =   360
            Left            =   1680
            TabIndex        =   4
            ToolTipText     =   "Ingrese fecha en formado dd/mm/yyyy"
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mebFecFin 
            Height          =   360
            Left            =   7440
            TabIndex        =   5
            ToolTipText     =   "Ingrese fecha en formado dd/mm/yyyy"
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo DatVendedor 
            Height          =   360
            Left            =   1680
            TabIndex        =   6
            Top             =   1320
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblActivo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4920
            TabIndex        =   26
            Tag             =   "X"
            Top             =   360
            Width           =   1995
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            Height          =   240
            Left            =   555
            TabIndex        =   20
            Top             =   1380
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio:"
            Height          =   240
            Left            =   285
            TabIndex        =   19
            Top             =   1860
            Width           =   1290
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fin:"
            Height          =   240
            Left            =   6240
            TabIndex        =   18
            Top             =   1800
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion:"
            Height          =   240
            Left            =   360
            TabIndex        =   17
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblIdMeta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1680
            TabIndex        =   16
            Tag             =   "X"
            Top             =   360
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id Meta:"
            Height          =   240
            Left            =   720
            TabIndex        =   15
            Top             =   397
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView lvDatos 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   2
         Top             =   1080
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12726
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   -73440
         TabIndex        =   1
         Top             =   600
         Width           =   9975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
         Height          =   240
         Left            =   -74640
         TabIndex        =   13
         Top             =   667
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private pIDempresa As Integer

Sub Mandar_Datos()
    cargarDatosAdicionales
    MousePointer = vbHourglass
    metaInfo pIDempresa, Me.lvDatos.SelectedItem.Tag

    With Me.lvDatos
    
        Me.lblIdMeta.Caption = .SelectedItem.Tag
        Me.txtDescripcion.Text = .SelectedItem.Text
        Me.DatVendedor.BoundText = .SelectedItem.SubItems(5)
        Me.mebFecIni.Text = .SelectedItem.SubItems(2)
        Me.mebFecFin.Text = .SelectedItem.SubItems(3)
        Me.lblActivo.Caption = .SelectedItem.SubItems(4)
    
        Estado_Botones AntesDeActualizar

    End With

    MousePointer = vbDefault

End Sub

Private Sub metaInfo(xIDempresa As Integer, xIDmeta As Integer)

    On Error GoTo xInfo
    Me.lvDetalle.ListItems.Clear

    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[dbo].[USP_META_INFO]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , xIDempresa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMETA", adInteger, adParamInput, , xIDmeta)
    
    Set oRSmain = oCmdEjec.Execute

    If Not oRSmain.EOF Then

        Dim itemx As Object

        Do While Not oRSmain.EOF
            Set itemx = Me.lvDetalle.ListItems.Add(, , oRSmain!prod, Me.ilMeta.ListImages(8).Key, Me.ilMeta.ListImages(8).Key)
            itemx.Tag = oRSmain!idp
            itemx.SubItems(1) = oRSmain!cnt
            itemx.SubItems(2) = oRSmain!imp
            oRSmain.MoveNext
        Loop
    
    End If
CerrarConexion True
    Exit Sub
xInfo:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub metaSearch(xdato As String)
    MousePointer = vbHourglass

    On Error GoTo xSearch

    Me.lvDatos.ListItems.Clear
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[dbo].[USP_META_SEARCH]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)

    If Len(Trim(xdato)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 100, xdato)
    
    Set oRSmain = oCmdEjec.Execute

    If Not oRSmain.EOF Then

        Dim itemx As Object

        Do While Not oRSmain.EOF
            Set itemx = Me.lvDatos.ListItems.Add(, , oRSmain!descr, Me.ilMeta.ListImages(8).Key, Me.ilMeta.ListImages(8).Key)
            itemx.Tag = oRSmain!ide
            itemx.SubItems(1) = oRSmain!vend
            itemx.SubItems(2) = oRSmain!FINI
            itemx.SubItems(3) = oRSmain!ffin
            itemx.SubItems(4) = oRSmain!ACT
            itemx.SubItems(5) = oRSmain!idv

            If oRSmain!ACT = "NO" Then
                Me.lvDatos.ListItems(itemx.Index).ForeColor = vbRed
                Me.lvDatos.ListItems(itemx.Index).ListSubItems(1).ForeColor = vbRed
                Me.lvDatos.ListItems(itemx.Index).ListSubItems(2).ForeColor = vbRed
                Me.lvDatos.ListItems(itemx.Index).ListSubItems(3).ForeColor = vbRed
                Me.lvDatos.ListItems(itemx.Index).ListSubItems(4).ForeColor = vbRed

            End If

            oRSmain.MoveNext
        Loop

    End If

    MousePointer = vbDefault
    CerrarConexion True
    Exit Sub
xSearch:
    MousePointer = vbDefault
    CerrarConexion True
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Estado_Botones(val As Valores)

    Select Case val

        Case InicializarFormulario, grabar, cancelar, Eliminar, Desactivar, Activar
            Me.mtbMeta.Buttons(1).Enabled = True
            Me.mtbMeta.Buttons(2).Enabled = False
            Me.mtbMeta.Buttons(3).Enabled = False
            Me.mtbMeta.Buttons(4).Enabled = False
            Me.mtbMeta.Buttons(5).Enabled = False
            Me.mtbMeta.Buttons(6).Enabled = False
            Me.mtbMeta.Buttons(7).Enabled = False
            Me.SSTMeta.tab = 0

        Case Nuevo, Editar
            Me.lblActivo.Caption = "SI"
            Me.mtbMeta.Buttons(1).Enabled = False
            Me.mtbMeta.Buttons(2).Enabled = True
            Me.mtbMeta.Buttons(3).Enabled = False
            Me.mtbMeta.Buttons(4).Enabled = True
            Me.mtbMeta.Buttons(5).Enabled = False
            Me.mtbMeta.Buttons(6).Enabled = False
            Me.mtbMeta.Buttons(7).Enabled = False
            Me.lvDatos.Enabled = False
            Me.txtSearch.Enabled = False
            Me.SSTMeta.tab = 1

        Case buscar
            Me.mtbMeta.Buttons(1).Enabled = True
            Me.mtbMeta.Buttons(2).Enabled = False
            Me.mtbMeta.Buttons(3).Enabled = False
            Me.mtbMeta.Buttons(4).Enabled = False
            Me.SSTMeta.tab = 0

        Case AntesDeActualizar
            Me.mtbMeta.Buttons(1).Enabled = False
            Me.mtbMeta.Buttons(2).Enabled = False
            Me.mtbMeta.Buttons(3).Enabled = True
            Me.mtbMeta.Buttons(4).Enabled = True

            If Me.lblActivo.Caption = "SI" Then
                Me.mtbMeta.Buttons(5).Enabled = True
                Me.mtbMeta.Buttons(6).Enabled = False

            Else
                Me.mtbMeta.Buttons(5).Enabled = False
                Me.mtbMeta.Buttons(6).Enabled = True
            
            End If
            Me.mtbMeta.Buttons(7).Enabled = True
            Me.SSTMeta.tab = 1

    End Select

End Sub

Private Sub cargarDatosAdicionales()

    On Error GoTo Adicional

    MousePointer = vbHourglass
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[dbo].[USP_META_DATOS_COMPLEMENTARIOS]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
     
    Set oRSmain = oCmdEjec.Execute
    
    ' Crear Recordsets temporales en memoria
    Dim orsTEMP1 As New ADODB.Recordset

    Dim orsTEMP2 As New ADODB.Recordset
 
    If Not oRSmain.EOF Then
    
        ' Configurar el primer Recordset temporal
        orsTEMP1.CursorLocation = adUseClient
        orsTEMP1.Fields.Append oRSmain(0).Name, oRSmain(0).Type, oRSmain(0).DefinedSize
        orsTEMP1.Fields.Append oRSmain(1).Name, oRSmain(1).Type, oRSmain(1).DefinedSize
        orsTEMP1.Open
    
        ' Copiar datos del primer Recordset
        oRSmain.MoveFirst

        Do Until oRSmain.EOF
            orsTEMP1.AddNew
            orsTEMP1.Fields(0).Value = oRSmain.Fields(0).Value
            orsTEMP1.Fields(1).Value = oRSmain.Fields(1).Value
            orsTEMP1.Update
            oRSmain.MoveNext
        Loop
    
        ' Configurar DataVendedor
        Set Me.DatVendedor.RowSource = orsTEMP1
        Me.DatVendedor.ListField = orsTEMP1.Fields(1).Name
        Me.DatVendedor.BoundColumn = orsTEMP1.Fields(0).Name
        Me.DatVendedor.BoundText = -1
    
        ' Obtener el segundo Recordset
        Set oRSmain = oRSmain.NextRecordset
    
        If Not oRSmain Is Nothing Then
            ' Configurar el segundo Recordset temporal
            orsTEMP2.CursorLocation = adUseClient
            orsTEMP2.Fields.Append oRSmain(0).Name, oRSmain(0).Type, oRSmain(0).DefinedSize
            orsTEMP2.Fields.Append oRSmain(1).Name, oRSmain(1).Type, oRSmain(1).DefinedSize
            orsTEMP2.Open
        
            ' Copiar datos del segundo Recordset
            If Not oRSmain.EOF Then
                oRSmain.MoveFirst

                Do Until oRSmain.EOF
                    orsTEMP2.AddNew
                    orsTEMP2.Fields(0).Value = oRSmain.Fields(0).Value
                    orsTEMP2.Fields(1).Value = oRSmain.Fields(1).Value
                    orsTEMP2.Update
                    oRSmain.MoveNext
                Loop

            End If
        
            ' Configurar datCategoria
            Set Me.DatProducto.RowSource = orsTEMP2
            Me.DatProducto.ListField = orsTEMP2.Fields(1).Name
            Me.DatProducto.BoundColumn = orsTEMP2.Fields(0).Name
            Me.DatProducto.BoundText = -1

        End If

    End If

    MousePointer = vbDefault
    CerrarConexion True

    Exit Sub
Adicional:
    MousePointer = vbDefault
    CerrarConexion True
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub cmdAdd_Click()

    If Me.DatProducto.BoundText = -1 Then
        MsgBox "Debe elegir el Producto.", vbInformation, Pub_Titulo
        Me.DatProducto.SetFocus
    ElseIf Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Debe ingresar la Cantidad.", vbInformation, Pub_Titulo
        Me.txtCantidad.SetFocus
    ElseIf Me.txtCantidad.Text <= 0 Then
        MsgBox "Cantidad ingresada incorrecta.", vbInformation, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
    ElseIf Len(Trim(Me.txtImporte.Text)) = 0 Then
        MsgBox "Debe ingresar el Importe.", vbInformation, Pub_Titulo
        Me.txtImporte.SetFocus
    ElseIf Me.txtImporte.Text <= 0 Then
        MsgBox "Importe ingresado incorrecto.", vbInformation, Pub_Titulo
        Me.txtImporte.SetFocus
        Me.txtImporte.SelStart = 0
        Me.txtImporte.SelLength = Len(Me.txtImporte.Text)
  
    Else

        Dim itemx As Object

        Dim xData As Boolean

        xData = False

        If Me.lvDetalle.ListItems.count = 0 Then

            Set itemx = Me.lvDetalle.ListItems.Add(, , Me.DatProducto.Text, Me.ilMeta.ListImages(8).Key, Me.ilMeta.ListImages(8).Key)
            itemx.Tag = Me.DatProducto.BoundText
            itemx.SubItems(1) = Me.txtCantidad.Text
            itemx.SubItems(2) = Me.txtImporte.Text
        Else

            For Each itemx In Me.lvDetalle.ListItems

                If itemx.Tag = Me.DatProducto.BoundText Then
                    xData = True
                    Exit For

                End If

            Next
        
            If xData Then
                MsgBox "Producto ya se encuentra en lista.", vbInformation, Pub_Titulo
                Me.DatProducto.SetFocus
                Exit Sub
            Else
                Set itemx = Me.lvDetalle.ListItems.Add(, , Me.DatProducto.Text, Me.ilMeta.ListImages(8).Key, Me.ilMeta.ListImages(8).Key)
                itemx.Tag = Me.DatProducto.BoundText
                itemx.SubItems(1) = Me.txtCantidad.Text
                itemx.SubItems(2) = Me.txtImporte.Text

            End If

        End If
        
        Me.DatProducto.BoundText = -1
        Me.txtCantidad.Text = ""
        Me.txtImporte.Text = ""
        
        Me.DatProducto.SetFocus

    End If

End Sub

Private Sub cmdDel_Click()

    If Me.lvDetalle.ListItems.count = 0 Then Exit Sub
    If Me.lvDetalle.SelectedItem Is Nothing Then Exit Sub

    
    Me.lvDetalle.ListItems.Remove Me.lvDetalle.SelectedItem.Index
End Sub

Private Sub DatProducto_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.txtCantidad
End Sub

Private Sub DatVendedor_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.mebFecIni
End Sub

Private Sub Form_Load()
pIDempresa = devuelveIDempresaXdefecto
configurarLV
Estado_Botones InicializarFormulario
  CentrarFormulario MDIForm1, Me
    Me.mtbMeta.ImageList = Me.ilMeta
    Dim i As Integer

    For i = 1 To 7
        Me.mtbMeta.Buttons(i).Image = Me.ilMeta.ListImages.Item(i).Index
    Next
End Sub

Private Sub configurarLV()
Me.lvDatos.Icons = Me.ilMeta
Me.lvDatos.SmallIcons = Me.ilMeta

Me.lvDetalle.Icons = Me.ilMeta
Me.lvDetalle.SmallIcons = Me.ilMeta

With Me.lvDetalle
    .ColumnHeaders.Add , , "Producto", 4500
    .ColumnHeaders.Add , , "Cantidad"
    .ColumnHeaders.Add , , "Importe"
    .HideColumnHeaders = False
    .View = lvwReport
    .FullRowSelect = True
End With

With Me.lvDatos

    .ColumnHeaders.Add , , "Descripcion", 4500
    .ColumnHeaders.Add , , "Vendedor", 3000
    .ColumnHeaders.Add , , "Fecha Ini"
    .ColumnHeaders.Add , , "Fecha Fin"
    .ColumnHeaders.Add , , "Activo"
    .ColumnHeaders.Add , , "idv", 0
    .HideColumnHeaders = False
    .View = lvwReport
.FullRowSelect = True
End With
End Sub

Private Sub lvDatos_DblClick()
DesactivarControles Me
Me.cmdAdd.Enabled = False
Me.cmdDel.Enabled = False
Mandar_Datos
End Sub

Private Sub mebFecFin_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.DatProducto
End Sub

Private Sub mebFecIni_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.mebFecFin
End Sub

Private Sub mtbMeta_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            cargarDatosAdicionales
            Me.lvDetalle.ListItems.Clear
            Estado_Botones Nuevo
            Me.cmdAdd.Enabled = True
            Me.cmdDel.Enabled = True
            VNuevo = True
            Me.txtDescripcion.SetFocus

        Case 2 'Guardar

            If Len(Trim(Me.txtDescripcion.Text)) = 0 Then
                MsgBox "Debe ingresar la Descripcion", vbCritical, Pub_Titulo
                Me.txtDescripcion.SetFocus
            ElseIf Me.DatVendedor.BoundText = -1 Then
                MsgBox "Debe elegir el Vendedor.", vbCritical, Pub_Titulo
                Me.DatVendedor.SetFocus
            ElseIf ValidarFecha(Me.mebFecIni.Text, True) = False Then
                MsgBox "Fecha Inicial incorrecta.", vbCritical, Pub_Titulo
                Me.mebFecIni.SetFocus
            ElseIf ValidarFecha(Me.mebFecFin.Text, True) = False Then
                MsgBox "Fecha Final incorrecta.", vbCritical, Pub_Titulo
                Me.mebFecFin.SetFocus
                '            ElseIf Me.DatVendedor.BoundText = -1 Then
                '                MsgBox "Debe elegir el vendedor.", vbInformation, Pub_Titulo
                '                Me.DatVendedor.SetFocus
            Else
                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True

                If VNuevo Then
                    oCmdEjec.CommandText = "[dbo].[USP_META_REGISTER]"
                Else
                    oCmdEjec.CommandText = "[dbo].[USP_META_UPDATE]"

                End If

                On Error GoTo grabar

                Dim Smensaje  As String

                Dim strFechaI As String, strFechaF As String

                strFechaI = Replace(Me.mebFecIni.Text, "_", "")
                strFechaF = Replace(Me.mebFecFin.Text, "_", "")
                
                strFechaI = ConvertirFechaFormat_yyyyMMdd(strFechaI)
                strFechaI = Replace(strFechaI, "/", "")
                
                strFechaF = ConvertirFechaFormat_yyyyMMdd(strFechaF)
                strFechaF = Replace(strFechaF, "/", "")

                Dim vIDz As Integer

                Smensaje = ""
                vIDz = 0

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, 2, pIDempresa)

                If Not VNuevo Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMETA", adInteger, adParamInput, , Me.lblIdMeta.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCRIPCION", adVarChar, adParamInput, 100, Trim(Me.txtDescripcion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAINI", adVarChar, adParamInput, 8, Trim(strFechaI))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAFIN", adVarChar, adParamInput, 8, Trim(strFechaF))
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTER", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Dim strItems As String

                strItems = ""

                Dim f As Integer
    
                If Me.lvDetalle.ListItems.count <> 0 Then
                    strItems = "<r>"

                    For f = 1 To Me.lvDetalle.ListItems.count
                        strItems = strItems & "<d "
                        strItems = strItems & "idp=""" & Me.lvDetalle.ListItems(f).Tag & """ "
                        strItems = strItems & "cnt=""" & Me.lvDetalle.ListItems(f).SubItems(1) & """ "
                        strItems = strItems & "imp=""" & Me.lvDetalle.ListItems(f).SubItems(2) & """ "
                        strItems = strItems & "/>"
                    Next
                    strItems = strItems & "</r>"

                End If
                
                If Len(Trim(strItems)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@XDETALLE", adVarChar, adParamInput, 4000, strItems)
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        DesactivarControles Me
                        Estado_Botones grabar
                        Me.lvDatos.Enabled = True
                        Me.txtSearch.Enabled = True
                        CerrarConexion True
                        'clienteSearch Me.txtSearch.Text
                        metaSearch Me.txtSearch.Text
                    Else
                        
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo
                        CerrarConexion True

                    End If

                End If

                MousePointer = vbDefault
                Exit Sub

grabar:
                MousePointer = vbDefault
                CerrarConexion True
                MsgBox Err.Description, vbInformation, Pub_Titulo

            End If

        Case 3 'Modificar
            VNuevo = False
            Estado_Botones Editar
            ActivarControles Me
            Me.cmdAdd.Enabled = True
            Me.cmdDel.Enabled = True
            Me.txtSearch.Enabled = False

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvDatos.Enabled = True
            Me.txtSearch.Enabled = True
            Me.txtSearch.SetFocus
            
        Case 5 'Desactivar
            
            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            
                On Error GoTo Desactiva

                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[dbo].[USP_META_STATE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMETA", adInteger, adParamInput, , Me.lblIdMeta.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STATE", adBoolean, adParamInput, , False)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        LimpiarControles Me
                        Estado_Botones Desactivar
                        Me.lvDatos.Enabled = True
                        metaSearch Me.txtSearch.Text
                    Else
                        CerrarConexion True
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If

                MousePointer = vbDefault
                Exit Sub
            
Desactiva:
                MousePointer = vbDefault
                CerrarConexion True
                MsgBox Err.Description, vbInformation, Pub_Titulo
            
            End If
            
        Case 6 'ACTIVAR
            
            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then

                On Error GoTo Activa

                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[dbo].[USP_META_STATE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMETA", adInteger, adParamInput, , Me.lblIdMeta.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STATE", adBoolean, adParamInput, , True)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        LimpiarControles Me
                        Estado_Botones Activar
                        Me.lvDatos.Enabled = True
                        metaSearch Me.txtSearch.Text
                    Else
                        CerrarConexion True
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If

                MousePointer = vbDefault
                Exit Sub
            
Activa:
                MousePointer = vbDefault
                CerrarConexion True
                MsgBox Err.Description, vbInformation, Pub_Titulo
            
            End If

        Case 7 'ELIMINAR

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            
                On Error GoTo Elimina

                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[dbo].[USP_META_DELETE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMETA", adInteger, adParamInput, , Me.lblIdMeta.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
              
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        DesactivarControles Me
                        Estado_Botones Eliminar
                        Me.lvDatos.Enabled = True
                        Me.txtSearch.Enabled = True
                
                        metaSearch Me.txtSearch.Text
                    Else
                        CerrarConexion True
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If

                MousePointer = vbDefault
                Exit Sub
            
Elimina:
                MousePointer = vbDefault
                CerrarConexion True
                MsgBox Err.Description, vbInformation, Pub_Titulo
            
            End If

    End Select

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
HandleEnterKey KeyAscii, Me.txtImporte
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.DatVendedor
End Sub

Private Sub txtImporte_Change()
ValidarSoloNumerosPunto Me.txtImporte
End Sub


Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumerosPunto(Me.txtImporte, KeyAscii)
 HandleEnterKey KeyAscii, Me.cmdAdd
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then metaSearch Me.txtSearch
End Sub
