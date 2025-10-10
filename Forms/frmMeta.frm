VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metas"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMeta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   16665
   Begin MSComctlLib.ImageList ilMeta 
      Left            =   17160
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":0CCA
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":1064
            Key             =   "save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":13FE
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":1798
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":1B32
            Key             =   "inactive"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":20CC
            Key             =   "active"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":2666
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":2A00
            Key             =   "meta"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMeta.frx":2D9A
            Key             =   "seller"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar mtbMeta 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   16665
      _ExtentX        =   29395
      _ExtentY        =   1111
      ButtonWidth     =   1720
      ButtonHeight    =   953
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
      Top             =   720
      Width           =   16455
      _ExtentX        =   29025
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
      TabPicture(0)   =   "frmMeta.frx":3134
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lvDatos"
      Tab(0).Control(1)=   "txtSearch"
      Tab(0).Control(2)=   "Label1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Meta"
      TabPicture(1)   =   "frmMeta.frx":3150
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FraDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraCabecera"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FraProducto"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame FraProducto 
         Height          =   6255
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   5415
         Begin VB.Frame fraMensajeVendedor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   240
            TabIndex        =   35
            Top             =   1900
            Width           =   4335
            Begin VB.Label lblMensajeVendedor 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Asigne un vendedor para comenzar."
               Height          =   240
               Left            =   0
               TabIndex        =   36
               Top             =   120
               Width           =   4275
            End
         End
         Begin VB.CommandButton cmdVendedorDel 
            Height          =   360
            Left            =   4800
            Picture         =   "frmMeta.frx":316C
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2040
            Width           =   510
         End
         Begin VB.CommandButton cmdVendedorAdd 
            Height          =   360
            Left            =   4800
            Picture         =   "frmMeta.frx":34F6
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1200
            Width           =   510
         End
         Begin MSDataListLib.DataCombo DatVendedor 
            Height          =   360
            Left            =   120
            TabIndex        =   26
            Top             =   1200
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvVendedor 
            Height          =   4335
            Left            =   120
            TabIndex        =   30
            Top             =   1800
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   7646
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
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "**2. Vendedores Asignados**"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   0
            TabIndex        =   28
            Top             =   120
            Width           =   5400
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Asignar Vendedor:"
            Height          =   240
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame FraCabecera 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   16215
         Begin VB.TextBox txtDescripcion 
            Height          =   375
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   3
            Tag             =   "X"
            Top             =   1200
            Width           =   7335
         End
         Begin MSMask.MaskEdBox mebFecIni 
            Height          =   375
            Left            =   9000
            TabIndex        =   4
            ToolTipText     =   "Ingrese fecha en formado dd/mm/yyyy"
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mebFecFin 
            Height          =   360
            Left            =   10800
            TabIndex        =   5
            ToolTipText     =   "Ingrese fecha en formado dd/mm/yyyy"
            Top             =   1200
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
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "**1. Definición y Periodo de la Meta**"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   0
            TabIndex        =   34
            Top             =   120
            Width           =   16200
         End
         Begin VB.Label lblActivo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6840
            TabIndex        =   24
            Tag             =   "X"
            Top             =   720
            Width           =   1995
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio"
            Height          =   240
            Left            =   9000
            TabIndex        =   18
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fin"
            Height          =   240
            Left            =   10800
            TabIndex        =   17
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion:"
            Height          =   240
            Left            =   240
            TabIndex        =   16
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblIdMeta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1560
            TabIndex        =   15
            Tag             =   "X"
            Top             =   720
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id Meta:"
            Height          =   240
            Left            =   600
            TabIndex        =   14
            Top             =   750
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView lvDatos 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   2
         Top             =   1080
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   12938
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
         Left            =   -73920
         TabIndex        =   1
         Top             =   600
         Width           =   15255
      End
      Begin VB.Frame FraDetalle 
         Height          =   6255
         Left            =   5640
         TabIndex        =   19
         Top             =   2160
         Width           =   10695
         Begin VB.Frame fraMensajeDetalle 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   360
            TabIndex        =   37
            Top             =   2040
            Width           =   8895
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Selecione un Vendedor para ver sus metas."
               Height          =   240
               Left            =   0
               TabIndex        =   38
               Top             =   120
               Width           =   8880
            End
         End
         Begin VB.CommandButton cmdDel 
            Height          =   360
            Left            =   9480
            Picture         =   "frmMeta.frx":3880
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2280
            Width           =   990
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   360
            Left            =   9480
            Picture         =   "frmMeta.frx":3C0A
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1920
            Width           =   990
         End
         Begin MSComctlLib.ListView lvDetalle 
            Height          =   4575
            Left            =   240
            TabIndex        =   9
            Top             =   1560
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   8070
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
            Left            =   6600
            TabIndex        =   8
            Tag             =   "X"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtCantidad 
            Height          =   375
            Left            =   1560
            TabIndex        =   7
            Tag             =   "X"
            Top             =   1080
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DatProducto 
            Height          =   360
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblIdVendedor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   9480
            TabIndex        =   33
            Top             =   3600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblMensajeProducto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "**3. Detalle de Productos/Metas para: **"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   10680
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
            Height          =   240
            Left            =   5640
            TabIndex        =   22
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
            Height          =   240
            Left            =   495
            TabIndex        =   21
            Top             =   1140
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Producto:"
            Height          =   240
            Left            =   480
            TabIndex        =   20
            Top             =   780
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
         Height          =   240
         Left            =   -74640
         TabIndex        =   12
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
Private oRSVendedor As ADODB.Recordset
Private oRSProducto As ADODB.Recordset

Private Function GenerarXML(ByVal oRSVendedor As ADODB.Recordset, _
                            ByVal oRSProducto As ADODB.Recordset) As String

    Dim sXML   As String

    Dim idVend As Long
    
    sXML = "<root>" & vbCrLf
    
    ' Recorrer vendedores
    oRSVendedor.Filter = ""
    oRSProducto.Filter = ""
    If Not oRSVendedor.EOF Then oRSVendedor.MoveFirst
    If Not oRSProducto.EOF Then oRSProducto.MoveFirst

    Do Until oRSVendedor.EOF
        idVend = oRSVendedor("idVendedor").Value
        
        sXML = sXML & "  <vendedor id=""" & idVend & """>" & vbCrLf
        
        ' Filtrar productos de ese vendedor
        If Not oRSProducto.EOF Then
            oRSProducto.MoveFirst

            Do Until oRSProducto.EOF

                If oRSProducto("idVendedor").Value = idVend Then
                    sXML = sXML & "    <detalle idProducto=""" & oRSProducto("idProducto").Value & """ cantidad=""" & oRSProducto("cantidad").Value & """ importe=""" & oRSProducto("importe").Value & """ />" & vbCrLf

                End If

                oRSProducto.MoveNext
            Loop
            oRSProducto.MoveFirst
        End If
        
        sXML = sXML & "  </vendedor>" & vbCrLf
        
        oRSVendedor.MoveNext
    Loop
    
    sXML = sXML & "</root>"
    
    ' Mostrar el resultado (puedes guardarlo en archivo o pasarlo como string)
    GenerarXML = sXML

End Function

Sub Mandar_Datos()
    cargarDatosAdicionales
    MousePointer = vbHourglass
    metaInfo pIDempresa, Me.lvDatos.SelectedItem.Tag

    With Me.lvDatos
    
        Me.lblIdMeta.Caption = .SelectedItem.Tag
        Me.txtDescripcion.Text = .SelectedItem.Text
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
    Me.lvVendedor.ListItems.Clear

    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[dbo].[USP_META_INFO]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , xIDempresa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMETA", adInteger, adParamInput, , xIDmeta)
    
    Set oRSmain = oCmdEjec.Execute

    Dim orsTmp As ADODB.Recordset
    
    LimpiarRecordsets
    
    If oRSmain.RecordCount <> 0 Then Me.fraMensajeVendedor.Visible = False
    
    Do While Not oRSmain.EOF
        agregarVendedorRS oRSmain!idv, oRSmain!vend
        agregarVendedorLV oRSmain!idv, oRSmain!vend
        oRSmain.MoveNext
    Loop
    
    Set orsTmp = oRSmain.NextRecordset
    
    Do While Not orsTmp.EOF
        agregarProductoRS orsTmp!idv, orsTmp!idp, orsTmp!prod, orsTmp!cant, orsTmp!imp
        orsTmp.MoveNext
    Loop

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
    ElseIf Len(Trim(Me.lblIdVendedor.Caption)) = 0 Then
        MsgBox "Debe elegir el vendedor.", vbInformation, Pub_Titulo
  
    Else

        Dim xData As Boolean

        xData = False
        
        oRSProducto.Filter = ""

        If Not oRSProducto.EOF Then oRSProducto.MoveFirst
        
        oRSProducto.Filter = "idVendedor=" & Me.lblIdVendedor.Caption & " and idProducto=" & Me.DatProducto.BoundText
        Me.fraMensajeDetalle.Visible = False

        If oRSProducto.EOF Then
            agregarProductoRS Me.lblIdVendedor.Caption, Me.DatProducto.BoundText, Me.DatProducto.Text, Me.txtCantidad.Text, Me.txtImporte.Text
            agregarProductoLV Me.lblIdVendedor.Caption, Me.DatProducto.BoundText, Me.DatProducto.Text, Me.txtCantidad.Text, Me.txtImporte.Text
            Me.DatProducto.BoundText = -1
            Me.txtCantidad.Text = ""
            Me.txtImporte.Text = ""
        
            Me.DatProducto.SetFocus
        Else
            MsgBox "Producto ya se encuentra en lista.", vbInformation, Pub_Titulo
            Me.DatProducto.SetFocus

        End If

    End If

End Sub

Private Sub cmdDel_Click()

    If Me.lvDetalle.ListItems.count = 0 Then Exit Sub
    If Me.lvDetalle.SelectedItem Is Nothing Then Exit Sub

    EliminarProducto Me.lblIdVendedor.Caption, Me.lvDetalle.SelectedItem.Tag
    
    Me.lvDetalle.ListItems.Remove Me.lvDetalle.SelectedItem.Index
    
    If Me.lvDetalle.ListItems.count = 0 Then Me.fraMensajeDetalle.Visible = True
End Sub

Private Sub cmdVendedorAdd_Click()

    If Me.DatVendedor.BoundText = -1 Then
        MsgBox "Debe elegir el vendedor.", vbInformation, Pub_Titulo
        Me.DatVendedor.SetFocus
    Else
        Me.fraMensajeVendedor.Visible = False

        Dim xData As Boolean

        xData = False
        
        oRSVendedor.Filter = ""

        If oRSVendedor.RecordCount <> 0 Then oRSVendedor.MoveFirst
        
        oRSVendedor.Filter = "idVendedor=" & Me.DatVendedor.BoundText
        
        If oRSVendedor.EOF Then
            agregarVendedorRS Me.DatVendedor.BoundText, Me.DatVendedor.Text
            agregarVendedorLV Me.DatVendedor.BoundText, Me.DatVendedor.Text
        Else
            MsgBox "Vendedor ya se encuentra en lista.", vbInformation, Pub_Titulo
            Me.DatVendedor.SetFocus
            Exit Sub

        End If
        
        oRSVendedor.Filter = ""

        If oRSVendedor.RecordCount <> 0 Then oRSVendedor.MoveFirst
        
    End If

End Sub

Private Sub agregarVendedorLV(cIDvendedor As Integer, cvendedor As String)

    Dim itemx As Object

    Set itemx = Me.lvVendedor.ListItems.Add(, , cvendedor, Me.ilMeta.ListImages(9).Key, Me.ilMeta.ListImages(9).Key)
    itemx.Tag = cIDvendedor

End Sub

Private Sub agregarProductoLV(cidVenvedor As Integer, _
                              cIDProducto As Integer, _
                              cProducto As String, _
                              cCantidad As Double, _
                              cImporte As Double)

    Dim itemx As Object

    Set itemx = Me.lvDetalle.ListItems.Add(, , cProducto, Me.ilMeta.ListImages(8).Key, Me.ilMeta.ListImages(8).Key)
    itemx.Tag = cIDProducto
    itemx.SubItems(1) = cCantidad
    itemx.SubItems(2) = cImporte

End Sub

Private Sub agregarVendedorRS(cIDvendedor As Integer, cvendedor As String)
oRSVendedor.AddNew
oRSVendedor("idVendedor").Value = cIDvendedor
oRSVendedor("vendedor").Value = cvendedor
oRSVendedor.Update
End Sub
Private Sub agregarProductoRS(cidVenvedor As Integer, cIDProducto As Integer, cProducto As String, cCantidad As Double, cImporte As Double)
oRSProducto.AddNew
oRSProducto("idVendedor").Value = cidVenvedor
oRSProducto("idProducto").Value = cIDProducto
oRSProducto("producto").Value = cProducto
oRSProducto("cantidad").Value = cCantidad
oRSProducto("importe").Value = cImporte
oRSProducto.Update
End Sub

Private Sub cmdVendedorDel_Click()

    If Me.lvVendedor.ListItems.count = 0 Then Exit Sub
    If Me.lvVendedor.SelectedItem Is Nothing Then Exit Sub
    EliminarVendedor Me.lvVendedor.SelectedItem.Tag
    Me.lvVendedor.ListItems.Remove Me.lvVendedor.SelectedItem.Index
    
    If Me.lvVendedor.ListItems.count = 0 Then Me.fraMensajeVendedor.Visible = True

End Sub

Private Sub EliminarProducto(ByVal idVend As Long, ByVal idProd As Long)
    If oRSProducto Is Nothing Then Exit Sub
    If oRSProducto.RecordCount = 0 Then Exit Sub
    
    oRSProducto.MoveFirst
    Do Until oRSProducto.EOF
        If oRSProducto!idvendedor = idVend And oRSProducto!idproducto = idProd Then
            oRSProducto.Delete
            Exit Do   ' salir porque ya eliminamos el producto
        End If
        oRSProducto.MoveNext
    Loop
    
    ' Opcional: reposicionar en el primer registro si aún quedan
    If oRSProducto.RecordCount > 0 Then
        oRSProducto.MoveFirst
    End If
End Sub


Private Sub EliminarVendedor(ByVal idBuscado As Long)
    Dim idVend As Long
    
    If oRSVendedor Is Nothing Or oRSProducto Is Nothing Then Exit Sub
    If oRSVendedor.RecordCount = 0 Then Exit Sub
    If oRSVendedor.RecordCount <> 0 Then oRSVendedor.MoveFirst
    If oRSProducto.RecordCount <> 0 Then oRSProducto.MoveFirst
    '--- Buscar vendedor

    Do Until oRSVendedor.EOF
        If oRSVendedor!idvendedor = idBuscado Then
            idVend = oRSVendedor!idvendedor
            
            ' 1. Eliminar vendedor
            oRSVendedor.Delete
            Exit Do
        End If
        oRSVendedor.MoveNext
    Loop
    
    '--- Si encontramos y borramos, eliminar también sus productos
    If idVend <> 0 Then
        If oRSProducto.RecordCount > 0 Then
            Do Until oRSProducto.EOF
                If oRSProducto!idvendedor = idVend Then
                    oRSProducto.Delete
                End If
                oRSProducto.MoveNext
            Loop
        End If
    End If
    
    ' Opcional: resetear cursor del vendedor
    If oRSVendedor.RecordCount > 0 Then
        oRSVendedor.MoveFirst
    End If
    Me.lvDetalle.ListItems.Clear
    Me.fraMensajeDetalle.Visible = True
End Sub


Private Sub DatProducto_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.txtCantidad
End Sub

Private Sub DatVendedor_KeyPress(KeyAscii As Integer)
cmdVendedorAdd_Click
End Sub

Private Sub Form_Load()
    pIDempresa = devuelveIDempresaXdefecto
    ConfigurarLV
    Estado_Botones InicializarFormulario
    CentrarFormulario MDIForm1, Me
    Me.mtbMeta.ImageList = Me.ilMeta

    Dim i As Integer

    For i = 1 To 7
        Me.mtbMeta.Buttons(i).Image = Me.ilMeta.ListImages.Item(i).Index
    Next
    configurarRecordsets

End Sub

Private Sub configurarRecordsets()
'CONFIGURAR RECORDSET VENDEDOR
Set oRSVendedor = New ADODB.Recordset

'definir campos
oRSVendedor.Fields.Append "idVendedor", adInteger
oRSVendedor.Fields.Append "vendedor", adVarChar, 100

'abrir recordset en memoria
oRSVendedor.CursorLocation = adUseClient
oRSVendedor.Open , , adOpenStatic, adLockBatchOptimistic

'CONFIGURAR RECORDSET PRODUCTO
Set oRSProducto = New ADODB.Recordset

'definir campos
oRSProducto.Fields.Append "idVendedor", adInteger
oRSProducto.Fields.Append "idProducto", adInteger
oRSProducto.Fields.Append "producto", adVarChar, 100
oRSProducto.Fields.Append "cantidad", adInteger
oRSProducto.Fields.Append "importe", adDouble

'abrir recordset en memoria
oRSProducto.CursorLocation = adUseClient
oRSProducto.Open , , adOpenStatic, adLockBatchOptimistic
End Sub

Private Sub ConfigurarLV()
    Me.lvDatos.Icons = Me.ilMeta
    Me.lvDatos.SmallIcons = Me.ilMeta

    Me.lvDetalle.Icons = Me.ilMeta
    Me.lvDetalle.SmallIcons = Me.ilMeta

    Me.lvVendedor.Icons = Me.ilMeta
    Me.lvVendedor.SmallIcons = Me.ilMeta

    With Me.lvDetalle
        .ColumnHeaders.Add , , "Producto", 4500
        .ColumnHeaders.Add , , "Cantidad"
        .ColumnHeaders.Add , , "Importe"
        .HideColumnHeaders = False
        .View = lvwReport
        .FullRowSelect = True

    End With

    With Me.lvDatos

        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Vendedores", 2000
        .ColumnHeaders.Add , , "Fecha Ini"
        .ColumnHeaders.Add , , "Fecha Fin"
        .ColumnHeaders.Add , , "Activo"
        .HideColumnHeaders = False
        .View = lvwReport
        .FullRowSelect = True

    End With

    With Me.lvVendedor

        .ColumnHeaders.Add , , "Vendedor", 3000
        .HideColumnHeaders = True
        .View = lvwReport
        .FullRowSelect = True

    End With

End Sub




Private Sub lvDatos_DblClick()
DesactivarControles Me
'Me.cmdAdd.Enabled = False
'Me.cmdDel.Enabled = False
Me.FraProducto.Enabled = False
Me.FraDetalle.Enabled = False
Mandar_Datos
End Sub

Private Sub muestraProductos(cIDvendedor As Integer)
    oRSProducto.Filter = ""

    If Not oRSProducto.EOF Then oRSProducto.MoveFirst
    oRSProducto.Filter = "idVendedor = " & cIDvendedor

    Me.lvDetalle.ListItems.Clear
    
    If oRSProducto.RecordCount <> 0 Then Me.fraMensajeDetalle.Visible = False

    Dim itemx As Object

    If Not oRSProducto.EOF Then
    Do While Not oRSProducto.EOF
        Set itemx = Me.lvDetalle.ListItems.Add(, , oRSProducto("Producto").Value, Me.ilMeta.ListImages(8).Key, Me.ilMeta.ListImages(8).Key)
        itemx.Tag = oRSProducto("idProducto").Value
        itemx.SubItems(1) = oRSProducto("cantidad").Value
        itemx.SubItems(2) = oRSProducto("importe").Value
        oRSProducto.MoveNext
    Loop
End If
End Sub

Private Sub lvVendedor_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.lblMensajeProducto.Caption = "**3. Detalle de Productos/Metas para: " & Me.lvVendedor.SelectedItem.Text & "**"
Me.lblIdVendedor.Caption = Me.lvVendedor.SelectedItem.Tag
muestraProductos Me.lvVendedor.SelectedItem.Tag
End Sub

Private Sub mebFecFin_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.DatVendedor
End Sub

Private Sub mebFecIni_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.mebFecFin
End Sub

Private Function validaProductosVendedores() As String
Dim sMsg As String, nomVend As String
Dim idVend As Integer
sMsg = ""
    If Not oRSVendedor.EOF Then oRSVendedor.MoveFirst
    
    oRSProducto.Filter = ""
    If oRSProducto.RecordCount <> 0 Then oRSProducto.MoveFirst
  

    Do Until oRSVendedor.EOF
        idVend = oRSVendedor("idVendedor").Value
        nomVend = oRSVendedor("Vendedor").Value
        bTieneProducto = False
        
        ' Buscar si este vendedor tiene productos
        If Not oRSProducto.EOF Then
            oRSProducto.MoveFirst

            Do Until oRSProducto.EOF

                If oRSProducto("idVendedor").Value = idVend Then
                    bTieneProducto = True
                    Exit Do

                End If

                oRSProducto.MoveNext
            Loop

        End If
        
        ' Si no tiene productos, lo acumulamos en mensaje
        If bTieneProducto = False Then
            sMsg = sMsg & vbCrLf & "El vendedor " & nomVend & " no tiene productos asignados."

        End If
        
        oRSVendedor.MoveNext
        If Not oRSProducto.EOF Then oRSProducto.MoveFirst
    Loop
    If oRSVendedor.RecordCount <> 0 Then oRSVendedor.MoveFirst
    If Not oRSProducto.EOF Then oRSProducto.MoveFirst
validaProductosVendedores = sMsg
End Function

Private Sub mtbMeta_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim xVend As String

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            cargarDatosAdicionales
            Me.lvDetalle.ListItems.Clear
            Estado_Botones Nuevo
            Me.FraProducto.Enabled = True
            Me.FraDetalle.Enabled = True
            VNuevo = True
            LimpiarRecordsets
            Me.lblIdVendedor.Caption = ""
            Me.lvVendedor.ListItems.Clear
            Me.txtDescripcion.SetFocus

        Case 2 'Guardar
          
            xVend = validaProductosVendedores

            If Len(Trim(Me.txtDescripcion.Text)) = 0 Then
                MsgBox "Debe ingresar la Descripcion", vbCritical, Pub_Titulo
                Me.txtDescripcion.SetFocus
            ElseIf ValidarFecha(Me.mebFecIni.Text, False) = False Then
                MsgBox "Fecha Inicial incorrecta.", vbCritical, Pub_Titulo
                Me.mebFecIni.SetFocus
            ElseIf ValidarFecha(Me.mebFecFin.Text, False) = False Then
                MsgBox "Fecha Final incorrecta.", vbCritical, Pub_Titulo
                Me.mebFecFin.SetFocus
                '            ElseIf Me.DatVendedor.BoundText = -1 Then
                '                MsgBox "Debe elegir el vendedor.", vbInformation, Pub_Titulo
                '                Me.DatVendedor.SetFocus
            ElseIf Me.lvVendedor.ListItems.count = 0 Then
                MsgBox "Debe agregar Vendedores para la Meta.", vbCritical, Pub_Titulo
            ElseIf Len(xVend) > 0 Then
                MsgBox "No puede proseguir debido a que: " & xVend
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
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAINI", adVarChar, adParamInput, 8, Trim(strFechaI))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAFIN", adVarChar, adParamInput, 8, Trim(strFechaF))
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTER", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Dim strItems As String
                
                strItems = GenerarXML(oRSVendedor, oRSProducto)
              
                If Len(Trim(strItems)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@XDETALLE", adLongVarWChar, adParamInput, Len(strItems), strItems)
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        'DesactivarControles Me
                        'Estado_Botones grabar
                        'Me.lvDatos.Enabled = True
                        'Me.txtSearch.Enabled = True
                       
                        'metaSearch Me.txtSearch.Text
                        If VNuevo Then Me.lblIdMeta.Caption = oRSmain!idmeta
                        
                        VNuevo = False
                        MsgBox oRSmain!Message, vbInformation, Pub_Titulo
                        CerrarConexion True
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
            Me.FraProducto.Enabled = True
            Me.FraDetalle.Enabled = True
            Me.txtSearch.Enabled = False

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            LimpiarRecordsets
            Me.lvDatos.Enabled = True
            Me.txtSearch.Enabled = True
            Me.txtSearch.SetFocus
            Me.lblMensajeProducto.Caption = "**3. Detalle de Productos/Metas para: **"
            metaSearch Me.txtSearch.Text
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
                        
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo
                        CerrarConexion True

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
                        DesactivarControles Me
                        Estado_Botones Eliminar
                        Me.lvDatos.Enabled = True
                        Me.txtSearch.Enabled = True
                        metaSearch Me.txtSearch.Text
                    Else
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                    CerrarConexion True

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
KeyAscii = Mayusculas(KeyAscii)
HandleEnterKey KeyAscii, Me.mebFecIni
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

Private Sub LimpiarRecordsets()
    oRSVendedor.Filter = ""
    oRSProducto.Filter = ""
    
If Not oRSVendedor.EOF Then oRSVendedor.MoveFirst
    If Not oRSProducto.EOF Then oRSProducto.MoveFirst
    If Not oRSVendedor Is Nothing Then
        If oRSVendedor.RecordCount > 0 Then
            oRSVendedor.MoveFirst
            Do Until oRSVendedor.EOF
                oRSVendedor.Delete
                oRSVendedor.MoveNext
            Loop
            'oRSVendedor.MoveFirst
        End If
    End If
    
    If Not oRSProducto Is Nothing Then
        If oRSProducto.RecordCount > 0 Then
            oRSProducto.MoveFirst
            Do Until oRSProducto.EOF
                oRSProducto.Delete
                oRSProducto.MoveNext
            Loop
            'oRSProducto.MoveFirst
        End If
    End If

End Sub

