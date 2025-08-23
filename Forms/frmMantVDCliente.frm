VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{FEC367D0-B73E-4DD0-80FD-1F56BC27B04A}#1.0#0"; "McToolBar.ocx"
Begin VB.Form frmMantVDCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes [Venta Directa]"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMantVDCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9495
   Begin ToolBar.McToolBar mtbCliente 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   7
      ButtonsWidth    =   90
      ButtonsHeight   =   70
      ButtonsPerRow   =   7
      HoverColor      =   -2147483635
      TooTipStyle     =   0
      ButtonsMode     =   4
      ButtonsPerRow_Chev=   7
      ButtonCaption1  =   "&Nuevo"
      ButtonIcon1     =   "frmMantVDCliente.frx":08CA
      ButtonToolTipIcon1=   1
      ButtonIconAllignment1=   0
      ButtonCaption2  =   "&Guardar"
      ButtonIcon2     =   "frmMantVDCliente.frx":15A4
      ButtonToolTipIcon2=   1
      ButtonIconAllignment2=   0
      ButtonCaption3  =   "&Modificar"
      ButtonIcon3     =   "frmMantVDCliente.frx":227E
      ButtonToolTipIcon3=   1
      ButtonIconAllignment3=   0
      ButtonCaption4  =   "&Cancelar"
      ButtonIcon4     =   "frmMantVDCliente.frx":2F58
      ButtonToolTipIcon4=   1
      ButtonIconAllignment4=   0
      ButtonCaption5  =   "&Desactivar"
      ButtonIcon5     =   "frmMantVDCliente.frx":3C32
      ButtonToolTipIcon5=   1
      ButtonIconAllignment5=   0
      ButtonCaption6  =   "&Activar"
      ButtonIcon6     =   "frmMantVDCliente.frx":490C
      ButtonToolTipIcon6=   1
      ButtonIconAllignment6=   0
      ButtonCaption7  =   "&Eliminar"
      ButtonIcon7     =   "frmMantVDCliente.frx":55E6
      ButtonToolTipIcon7=   1
      ButtonIconAllignment7=   0
   End
   Begin MSComctlLib.ImageList ilCliente 
      Left            =   9720
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVDCliente.frx":62C0
            Key             =   "client"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTCliente 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12515
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
      TabPicture(0)   =   "frmMantVDCliente.frx":685A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtSearch"
      Tab(0).Control(1)=   "lvCliente"
      Tab(0).Control(2)=   "Label1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Cliente"
      TabPicture(1)   =   "frmMantVDCliente.frx":6876
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblIdCliente"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FraRuc"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FraDni"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "optRuc"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "optDni"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "FraDatospersonales"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.Frame FraDatospersonales 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   720
         TabIndex        =   19
         Top             =   2400
         Width           =   7695
         Begin MSMask.MaskEdBox mebFecNac 
            Height          =   360
            Left            =   2040
            TabIndex        =   9
            ToolTipText     =   "Ingrese fecha en formado dd/mm/yyyy"
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
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
            Left            =   2040
            TabIndex        =   12
            Top             =   2280
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtTelefono 
            Height          =   360
            Left            =   5280
            TabIndex        =   10
            Tag             =   "X"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtNombreCorto 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            TabIndex        =   8
            Tag             =   "X"
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox txtDireccion 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            TabIndex        =   11
            Tag             =   "X"
            Top             =   1800
            Width           =   5055
         End
         Begin VB.TextBox txtRS 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            TabIndex        =   7
            Tag             =   "X"
            Top             =   360
            Width           =   5055
         End
         Begin MSDataListLib.DataCombo datCategoria 
            Height          =   360
            Left            =   2040
            TabIndex        =   13
            Top             =   2760
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Activo:"
            Height          =   240
            Left            =   1125
            TabIndex        =   28
            Top             =   3300
            Width           =   720
         End
         Begin VB.Label lblActivo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   2040
            TabIndex        =   27
            Tag             =   "X"
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nac.:"
            Height          =   240
            Left            =   630
            TabIndex        =   26
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono:"
            Height          =   240
            Left            =   4320
            TabIndex        =   25
            Top             =   1380
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria:"
            Height          =   240
            Left            =   810
            TabIndex        =   24
            Top             =   2820
            Width           =   1035
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            Height          =   240
            Left            =   825
            TabIndex        =   23
            Top             =   2340
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Corto:"
            Height          =   240
            Left            =   405
            TabIndex        =   22
            Top             =   900
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   240
            Left            =   855
            TabIndex        =   21
            Top             =   1860
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   240
            Left            =   510
            TabIndex        =   20
            Top             =   420
            Width           =   1335
         End
      End
      Begin VB.OptionButton optDni 
         Caption         =   "Dni"
         Height          =   255
         Left            =   4920
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton optRuc 
         Caption         =   "Ruc"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
      End
      Begin VB.Frame FraDni 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   18
         Top             =   1320
         Width           =   3735
         Begin VB.TextBox txtdni 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            MaxLength       =   8
            TabIndex        =   4
            Tag             =   "X"
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame FraRuc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   17
         Top             =   1320
         Width           =   3735
         Begin VB.TextBox txtRuc 
            Height          =   360
            Left            =   240
            MaxLength       =   11
            TabIndex        =   3
            Tag             =   "X"
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.TextBox txtSearch 
         Height          =   360
         Left            =   -73560
         TabIndex        =   1
         Top             =   480
         Width           =   7695
      End
      Begin MSComctlLib.ListView lvCliente 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   2
         Top             =   960
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
      Begin VB.Label lblIdCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   2280
         TabIndex        =   16
         Tag             =   "X"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id Cliente:"
         Height          =   240
         Left            =   1200
         TabIndex        =   15
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Busqueda"
         Height          =   240
         Left            =   -74640
         TabIndex        =   14
         Top             =   540
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmMantVDCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private pIDempresa As Integer

Private Sub cargarDatosAdicionales()

    On Error GoTo Adicional

    MousePointer = vbHourglass
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_CLIENTE_DATOS_COMPLEMENTARIOS]"
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
            Set Me.datCategoria.RowSource = orsTEMP2
            Me.datCategoria.ListField = orsTEMP2.Fields(1).Name
            Me.datCategoria.BoundColumn = orsTEMP2.Fields(0).Name
            Me.datCategoria.BoundText = -1

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

Sub Mandar_Datos()
    MousePointer = vbHourglass

    With Me.lvCliente
        Me.lblIdCliente.Caption = .SelectedItem.Text
        Me.txtRS.Text = .SelectedItem.SubItems(1)
        Me.lblActivo.Caption = .SelectedItem.SubItems(2)
        
        LimpiaParametros oCmdEjec, True
        oCmdEjec.CommandText = "[vd].[USP_CLIENTE_DATOS_ADICIONALES]"
        oCmdEjec.CommandType = adCmdStoredProc
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.lblIdCliente.Caption)
        oCmdEjec.Execute
        
        Set oRSmain = oCmdEjec.Execute
        
        If Not oRSmain.EOF Then
        
            If Len(Trim(oRSmain!RUC)) <> 0 Then
                Me.optRuc.Value = True
            Else
                Me.optDni.Value = True

            End If

            Me.txtRuc.Text = oRSmain!RUC
            Me.txtdni.Text = oRSmain!DNI
            Me.txtDireccion.Text = oRSmain!dir
            Me.DatVendedor.BoundText = oRSmain!IDVEN
            Me.datCategoria.BoundText = oRSmain!IDcat

            If Len(Trim(oRSmain!FNAC)) <> 0 Then Me.mebFecNac.Text = oRSmain!FNAC
            Me.txtTelefono.Text = oRSmain!TEL
            Me.txtNombreCorto.Text = oRSmain!NCORTO

        End If

        Me.txtRuc.Enabled = False
        Me.txtdni.Enabled = False
        CerrarConexion True
        Estado_Botones AntesDeActualizar

    End With

    MousePointer = vbDefault

End Sub

Private Sub clienteSearch(xdato As String)
MousePointer = vbHourglass
    On Error GoTo xSearch

    Me.lvCliente.ListItems.Clear
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_CLIENTE_SEARCH]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)

    If Len(Trim(xdato)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 100, xdato)
    
    Set oRSmain = oCmdEjec.Execute

    If Not oRSmain.EOF Then

        Dim itemx As Object

        Do While Not oRSmain.EOF
            Set itemx = Me.lvCliente.ListItems.Add(, , oRSmain!ide, Me.ilCliente.ListImages(1).Key, Me.ilCliente.ListImages(1).Key)
            itemx.SubItems(1) = oRSmain!rso
            itemx.SubItems(2) = oRSmain!ACT

            If oRSmain!ACT = "NO" Then
                Me.lvCliente.ListItems(itemx.Index).ForeColor = vbRed
                Me.lvCliente.ListItems(itemx.Index).ListSubItems(1).ForeColor = vbRed
                Me.lvCliente.ListItems(itemx.Index).ListSubItems(2).ForeColor = vbRed

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
            Me.mtbCliente.Button_Index = 1
            Me.mtbCliente.ButtonEnabled = True
            Me.mtbCliente.Button_Index = 2
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 3
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 4
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 5
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 6
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 7
            Me.mtbCliente.ButtonEnabled = False
            Me.SSTCliente.tab = 0

        Case Nuevo, Editar
            Me.lblActivo.Caption = "SI"
            Me.mtbCliente.Button_Index = 1
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 2
            Me.mtbCliente.ButtonEnabled = True
            Me.mtbCliente.Button_Index = 3
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 4
            Me.mtbCliente.ButtonEnabled = True
            Me.mtbCliente.Button_Index = 5
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 6
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 7
            Me.mtbCliente.ButtonEnabled = False
            Me.lvCliente.Enabled = False
            Me.txtSearch.Enabled = False
            Me.SSTCliente.tab = 1

        Case buscar
            Me.mtbCliente.Button_Index = 1
            Me.mtbCliente.ButtonEnabled = True
            Me.mtbCliente.Button_Index = 2
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 3
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 4
            Me.mtbCliente.ButtonEnabled = False
            Me.SSTCliente.tab = 0

        Case AntesDeActualizar
            Me.mtbCliente.Button_Index = 1
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 2
            Me.mtbCliente.ButtonEnabled = False
            Me.mtbCliente.Button_Index = 3
            Me.mtbCliente.ButtonEnabled = True
            Me.mtbCliente.Button_Index = 4
            Me.mtbCliente.ButtonEnabled = True

            If Me.lblActivo.Caption = "SI" Then
                Me.mtbCliente.Button_Index = 5
                Me.mtbCliente.ButtonEnabled = True
                Me.mtbCliente.Button_Index = 6
                Me.mtbCliente.ButtonEnabled = False

            Else
                Me.mtbCliente.Button_Index = 5
                Me.mtbCliente.ButtonEnabled = False
                Me.mtbCliente.Button_Index = 6
                Me.mtbCliente.ButtonEnabled = True
            
            End If

            Me.mtbCliente.Button_Index = 7
            Me.mtbCliente.ButtonEnabled = True
            Me.SSTCliente.tab = 1

    End Select

End Sub

Private Sub ConfigurarLV()
Me.lvCliente.Icons = Me.ilCliente
Me.lvCliente.SmallIcons = Me.ilCliente

With Me.lvCliente
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "IDE"
    .ColumnHeaders.Add , , "CATEGORIA", 3000
    .ColumnHeaders.Add , , "ACTIVO"
End With
End Sub

Private Sub DatVendedor_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.datCategoria
End Sub

Private Sub dtpFechNac_KeyDown(KeyCode As Integer, Shift As Integer)
   HandleEnterKey KeyCode, Me.txtTelefono
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
pIDempresa = devuelveIDempresaXdefecto
ConfigurarLV
DesactivarControles Me
Estado_Botones InicializarFormulario
clienteSearch Me.txtSearch.Text
CentrarFormulario MDIForm1, Me
End Sub

Private Sub lvCliente_DblClick()
cargarDatosAdicionales
Mandar_Datos
End Sub

Private Sub mebFecNac_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.txtTelefono
End Sub

Private Sub mtbCliente_Click(ByVal ButtonIndex As Long)

    Select Case ButtonIndex

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            cargarDatosAdicionales
            Estado_Botones Nuevo
            Me.optDni.Value = False
            Me.optRuc.Value = False
            Me.txtdni.Enabled = False
            Me.txtRuc.Enabled = False
            VNuevo = True
            Me.optRuc.Value = True

        Case 2 'Guardar

            If Len(Trim(Me.txtRS.Text)) = 0 Then
                MsgBox "Debe ingresar la Denominación", vbCritical, Pub_Titulo
                Me.txtRS.SetFocus
          
            ElseIf ValidarFecha(Me.mebFecNac.Text, True) = False Then
                MsgBox "Fecha incorrecta.", vbCritical, Pub_Titulo
                Me.mebFecNac.SetFocus
                ElseIf Me.DatVendedor.BoundText = -1 Then
                MsgBox "Debe elegir el vendedor.", vbInformation, Pub_Titulo
                Me.DatVendedor.SetFocus
            Else
                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True

                If VNuevo Then
                    oCmdEjec.CommandText = "[vd].[USP_CLIENTE_REGISTER]"
                Else
                    oCmdEjec.CommandText = "[vd].[USP_CLIENTE_UPDATE]"

                End If

                On Error GoTo grabar

                Dim Smensaje As String

                Dim strFecha As String

                strFecha = Replace(Me.mebFecNac.Text, "_", "")
                
                strFecha = ConvertirFechaFormat_yyyyMMdd(strFecha)
                strFecha = Replace(strFecha, "/", "")

                Dim vIDz As Integer

                Smensaje = ""
                vIDz = 0

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, 2, pIDempresa)

                If Not VNuevo Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.lblIdCliente.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, Trim(Me.txtRuc.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adVarChar, adParamInput, 8, Trim(Me.txtdni.Text))
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RS", adVarChar, adParamInput, 100, Trim(Me.txtRS.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIR", adVarChar, adParamInput, 300, Trim(Me.txtDireccion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCATEGORIA", adInteger, adParamInput, , Me.datCategoria.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECNAC", adVarChar, adParamInput, 8, Trim(strFecha))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TELEFONO", adVarChar, adParamInput, 10, Trim(Me.txtTelefono.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMCORTO", adVarChar, adParamInput, 20, Trim(Me.txtNombreCorto.Text))
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@UDUSARIO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        DesactivarControles Me
                        Estado_Botones grabar
                        Me.lvCliente.Enabled = True
                        Me.txtSearch.Enabled = True
                        CerrarConexion True
                        clienteSearch Me.txtSearch.Text
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

            If Me.optRuc.Value Then Me.txtRuc.SetFocus
            If Me.optDni.Value Then Me.txtdni.SetFocus
            Me.txtSearch.Enabled = False

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvCliente.Enabled = True
            Me.txtSearch.Enabled = True
            Me.txtSearch.SetFocus
            
        Case 5 'Desactivar
            
            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            
                On Error GoTo Desactiva

                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[vd].[USP_CLIENTE_STATE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.lblIdCliente.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STATE", adBoolean, adParamInput, , False)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        LimpiarControles Me
                        Estado_Botones Desactivar
                        Me.lvCliente.Enabled = True
                        clienteSearch Me.txtSearch.Text
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
                oCmdEjec.CommandText = "[vd].[USP_CLIENTE_STATE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.lblIdCliente.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STATE", adBoolean, adParamInput, , True)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        LimpiarControles Me
                        Estado_Botones Activar
                        Me.lvCliente.Enabled = True
                        clienteSearch Me.txtSearch.Text
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
                oCmdEjec.CommandText = "[vd].[USP_CLIENTE_DELETE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.lblIdCliente.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
              
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        DesactivarControles Me
                        Estado_Botones Eliminar
                        Me.lvCliente.Enabled = True
                        Me.txtSearch.Enabled = True
                
                        clienteSearch Me.txtSearch.Text
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

Private Sub optDni_Click()
If Me.optDni.Value Then
    Me.txtRuc.Text = ""
    Me.txtdni.Enabled = True
    Me.txtdni.SetFocus
End If

End Sub

Private Sub optRuc_Click()
If Me.optRuc.Value Then
    Me.txtdni.Text = ""
    Me.txtRuc.Enabled = True
    Me.txtRuc.SetFocus
End If
End Sub

Private Sub tbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
  

End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
HandleEnterKey KeyAscii, Me.DatVendedor
End Sub

Private Sub txtdni_Change()
ValidarSoloNumeros Me.txtdni
End Sub

Private Sub txtNombreCorto_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
HandleEnterKey KeyAscii, Me.mebFecNac
End Sub

Private Sub txtRS_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
HandleEnterKey KeyAscii, Me.txtNombreCorto
End Sub

Private Sub txtdni_KeyPress(KeyAscii As Integer)
 KeyAscii = SoloNumeros(KeyAscii)
HandleEnterKey KeyAscii, Me.txtRS
End Sub

Private Sub txtRuc_Change()
ValidarSoloNumeros Me.txtRuc
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
 HandleEnterKey KeyAscii, Me.txtRS
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then clienteSearch Me.txtSearch.Text
End Sub

Private Sub txtTelefono_Change()
ValidarSoloNumeros Me.txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
 KeyAscii = SoloNumeros(KeyAscii)
KeyAscii = Mayusculas(KeyAscii)
HandleEnterKey KeyAscii, Me.txtDireccion
End Sub
