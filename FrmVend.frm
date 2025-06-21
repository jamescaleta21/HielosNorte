VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FrmVen 
   Caption         =   "Maestro de Vendedores"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   ControlBox      =   0   'False
   Icon            =   "FrmVend.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11895
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmVend.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   1650
      Width           =   1185
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      Picture         =   "FrmVend.frx":0F2C
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   2640
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      Picture         =   "FrmVend.frx":1CEE
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   570
      Width           =   1185
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      Picture         =   "FrmVend.frx":2B88
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   4890
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      Picture         =   "FrmVend.frx":33FE
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   3810
      Width           =   1185
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame F1 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10095
      Begin VB.Frame FraMovil 
         Height          =   2775
         Left            =   6600
         TabIndex        =   83
         Top             =   120
         Width           =   3135
         Begin MSDataListLib.DataCombo DatEmpresa 
            Height          =   315
            Left            =   1080
            TabIndex        =   96
            Top             =   2400
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.ComboBox ComPerfil 
            Height          =   315
            ItemData        =   "FrmVend.frx":3BAC
            Left            =   1080
            List            =   "FrmVend.frx":3BBC
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   2040
            Width           =   1695
         End
         Begin VB.ComboBox comPrecio 
            Height          =   315
            ItemData        =   "FrmVend.frx":3BF7
            Left            =   1080
            List            =   "FrmVend.frx":3C01
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtPass 
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   91
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtUser 
            Height          =   285
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   90
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox comLogeo 
            Height          =   315
            ItemData        =   "FrmVend.frx":3C0D
            Left            =   1080
            List            =   "FrmVend.frx":3C17
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox ComActivo 
            Height          =   315
            ItemData        =   "FrmVend.frx":3C23
            Left            =   1080
            List            =   "FrmVend.frx":3C2D
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pertenece a:"
            Height          =   195
            Left            =   135
            TabIndex        =   97
            Top             =   2460
            Width           =   915
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Perfil:"
            Height          =   195
            Left            =   660
            TabIndex        =   94
            Top             =   2100
            Width           =   390
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edita Precios"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   1740
            Width           =   930
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   465
            TabIndex        =   89
            Top             =   1005
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Height          =   195
            Left            =   315
            TabIndex        =   88
            Top             =   1365
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Activo:"
            Height          =   195
            Left            =   555
            TabIndex        =   85
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logeo Movil:"
            Height          =   195
            Left            =   135
            TabIndex        =   84
            Top             =   660
            Width           =   915
         End
      End
      Begin VB.ComboBox cmbtransporte 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txttelecelu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   12
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txttelecasa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtdireccion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txtnombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox Txt_key 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtfechaing 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tranportista :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   25
         Left            =   3240
         TabIndex        =   76
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telf.Celular :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telf. Domicilio :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fec.  Ingreso :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   600
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3510
      Left            =   120
      TabIndex        =   17
      Top             =   3495
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6191
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Definición de Serie"
      TabPicture(0)   =   "FrmVend.frx":3C39
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "F2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Archivos de Impresión"
      TabPicture(1)   =   "FrmVend.frx":3C55
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label5(4)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtfac"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtbol"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtguia"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtgr"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtnc"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtnd"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Datos de Repartidor"
      TabPicture(2)   =   "FrmVend.frx":3C71
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(2)=   "Label16"
      Tab(2).Control(3)=   "txtBrevete"
      Tab(2).Control(4)=   "txtPlaca"
      Tab(2).Control(5)=   "txtCapacidad"
      Tab(2).ControlCount=   6
      Begin VB.TextBox txtCapacidad 
         Height          =   330
         Left            =   -72120
         TabIndex        =   103
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtPlaca 
         Height          =   330
         Left            =   -72120
         TabIndex        =   102
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtBrevete 
         Height          =   330
         Left            =   -72120
         TabIndex        =   101
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtnd 
         Height          =   285
         Left            =   1320
         TabIndex        =   74
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtnc 
         Height          =   285
         Left            =   1320
         TabIndex        =   72
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtgr 
         Height          =   285
         Left            =   1320
         TabIndex        =   70
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtguia 
         Height          =   285
         Left            =   1320
         TabIndex        =   68
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtbol 
         Height          =   285
         Left            =   1320
         TabIndex        =   66
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtfac 
         Height          =   285
         Left            =   1320
         TabIndex        =   64
         Top             =   720
         Width           =   2175
      End
      Begin VB.Frame F2 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   18
         Top             =   360
         Width           =   9615
         Begin VB.TextBox numfac_b 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            MaxLength       =   9
            TabIndex        =   43
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox Serie_b 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            MaxLength       =   4
            TabIndex        =   42
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox numfac_g 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   41
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox serie_g 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   40
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox serie_f 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            MaxLength       =   4
            TabIndex        =   39
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox numfac_f 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            MaxLength       =   9
            TabIndex        =   38
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox numfac_g_f 
            Height          =   285
            Left            =   1680
            TabIndex        =   37
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox numfac_b_f 
            Height          =   285
            Left            =   3240
            TabIndex        =   36
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox numfac_f_f 
            Height          =   285
            Left            =   4800
            TabIndex        =   35
            Top             =   2640
            Width           =   975
         End
         Begin VB.CheckBox cheguia 
            Alignment       =   1  'Right Justify
            Caption         =   "Inicializar - Serie Guia "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1680
            TabIndex        =   34
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox cheboleta 
            Alignment       =   1  'Right Justify
            Caption         =   "Inicializar - Serie Boleta "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   3240
            TabIndex        =   33
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chefactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Inicializar - Serie Factura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4800
            TabIndex        =   32
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Inicializar - Serie Ped"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox numfac_p_f 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox serie_p 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   29
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox numfac_p 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   9
            TabIndex        =   28
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox remi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   27
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chenc 
            Alignment       =   1  'Right Justify
            Caption         =   "Inicializar - Serie N/C"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   6360
            TabIndex        =   26
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox numfac_nc_f 
            Height          =   285
            Left            =   6390
            TabIndex        =   25
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox numfac_nc 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6390
            MaxLength       =   9
            TabIndex        =   24
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox serie_nc 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6390
            MaxLength       =   4
            TabIndex        =   23
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox chend 
            Alignment       =   1  'Right Justify
            Caption         =   "Inicializar - Serie N/D"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   7920
            TabIndex        =   22
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox numfac_nd_f 
            Height          =   285
            Left            =   7920
            TabIndex        =   21
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox numfac_nd 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7920
            MaxLength       =   9
            TabIndex        =   20
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox serie_nd 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7920
            MaxLength       =   4
            TabIndex        =   19
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. Boleta Inicial:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   62
            Top             =   1800
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serie Boleta :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   3240
            TabIndex        =   61
            Top             =   1200
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serie Guia :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   1680
            TabIndex        =   60
            Top             =   1200
            Width           =   855
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. Guia Inicial  :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   1680
            TabIndex        =   59
            Top             =   1800
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serie Factura :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   4800
            TabIndex        =   58
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N.Factura Inicial:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   4800
            TabIndex        =   57
            Top             =   1800
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N.Guia  Final :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   1680
            TabIndex        =   56
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. Boleta Final :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   3240
            TabIndex        =   55
            Top             =   2400
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. Factura Final :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   4800
            TabIndex        =   54
            Top             =   2400
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N.Ped. Final"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   53
            Top             =   2400
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N.Ped. Inicial "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   52
            Top             =   1800
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serie Pedido:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   51
            Top             =   1200
            Width           =   990
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "Serie de Guia de Remisión :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   18
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. N/C Final :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   6390
            TabIndex        =   49
            Top             =   2400
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro.N/C Inicial:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   6390
            TabIndex        =   48
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serie N/C :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   6390
            TabIndex        =   47
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro.N/D  Final :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   7920
            TabIndex        =   46
            Top             =   2415
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. N/D Inicial:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   7920
            TabIndex        =   45
            Top             =   1815
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serie N/D :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   7920
            TabIndex        =   44
            Top             =   1215
            Width           =   780
         End
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAPACIDAD EN KG:"
         Height          =   195
         Left            =   -73785
         TabIndex        =   100
         Top             =   1988
         Width           =   1500
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLACA:"
         Height          =   195
         Left            =   -72840
         TabIndex        =   99
         Top             =   1388
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BREVETE:"
         Height          =   195
         Left            =   -73080
         TabIndex        =   98
         Top             =   788
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nta Deb. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   75
         Top             =   2520
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nta Cred.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   73
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "G. Remisión :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   71
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Guia :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   69
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Boletas :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   67
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Facturas :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   65
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Archivos de Impresión:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   255
         TabIndex        =   63
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Timer PARPADEA 
      Interval        =   100
      Left            =   120
      Top             =   4800
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10320
      TabIndex        =   16
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "! Talonarios esta Definido por Compañia !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   135
      TabIndex        =   15
      Top             =   3465
      Width           =   4380
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   900
   End
End
Attribute VB_Name = "FrmVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pasa As Boolean
Dim loc_key As Integer
Dim CU As Integer

Dim PS_TRAONE As rdoQuery
Dim TRANSPORTEONE As rdoResultset

Public Function GENERA_VEN() As Integer
Dim valor As Integer
Dim ven_loc As rdoResultset
Dim PSVEN_LOC  As rdoQuery
pub_cadena = "SELECT VEM_CODVEN FROM VEMAEST WHERE VEM_CODCIA  = ?  ORDER BY VEM_CODVEN"
Set PSVEN_LOC = CN.CreateQuery("", pub_cadena)
PSVEN_LOC(0) = 0
Set ven_loc = PSVEN_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
PSVEN_LOC(0) = LK_CODCIA
ven_loc.Requery
If ven_loc.EOF Then
 valor = 0
Else
 ven_loc.MoveLast
 valor = ven_loc!VEM_CODVEN
End If
GENERA_VEN = valor + 1

End Function

Public Sub GRABAR_VEN()

    Dim NAMETRA As String

    If Left(cmdModificar.Caption, 2) = "&G" Then
        ven_llave.Edit
    Else
        ven_llave.AddNew
       
    End If

    ven_llave!VEM_CODVEN = val(FrmVen.Txt_key.Text)
    ven_llave!VEM_NOMBRE = FrmVen.txtnombre.Text
    ven_llave!vem_codcia = LK_CODCIA
    ven_llave!VEM_SERIE_G = val(FrmVen.serie_g.Text)
    ven_llave!VEM_NUMFAC_G_INI = val(FrmVen.numfac_g.Text)
    ven_llave!VEM_SERIE_B = val(FrmVen.Serie_b.Text)
    ven_llave!VEM_NUMFAC_B_INI = val(FrmVen.numfac_b.Text)
    ven_llave!VEM_SERIE_F = val(FrmVen.serie_f.Text)

    ven_llave!VEM_NUMFAC_F_INI = val(FrmVen.numfac_f.Text)
    ven_llave!VEM_NUMFAC_G_FIN = val(FrmVen.numfac_g_f.Text)
    ven_llave!VEM_NUMFAC_B_FIN = val(FrmVen.numfac_b_f.Text)
    ven_llave!VEM_NUMFAC_F_FIN = val(FrmVen.numfac_f_f.Text)

    ven_llave!VEM_SERIE_P = val(FrmVen.serie_p.Text)
    ven_llave!VEM_NUMFAC_P_INI = val(FrmVen.numfac_p.Text)
    ven_llave!VEM_NUMFAC_P_FIN = val(FrmVen.numfac_p_f.Text)
    ven_llave!VEM_FLAG_P = " "

    If Check1.Value = 1 Then
        ven_llave!VEM_FLAG_P = "A"

    End If

    ven_llave!VEM_SERIE_N = val(FrmVen.serie_nc.Text)
    ven_llave!VEM_SERIE_D = val(FrmVen.serie_nd.Text)
    ven_llave!VEM_NUMFAC_N_INI = val(FrmVen.numfac_nc.Text)
    ven_llave!VEM_NUMFAC_D_INI = val(FrmVen.numfac_nd.Text)
    ven_llave!VEM_NUMFAC_N_FIN = val(FrmVen.numfac_nc_f.Text)
    ven_llave!VEM_NUMFAC_D_FIN = val(FrmVen.numfac_nd_f.Text)
    ven_llave!VEM_FLAG_N = " "

    If chenc.Value = 1 Then
        ven_llave!VEM_FLAG_N = "A"

    End If

    ven_llave!VEM_FLAG_D = " "

    If chend.Value = 1 Then
        ven_llave!VEM_FLAG_D = "A"

    End If

    ven_llave!VEM_FECHA_ING = txtfechaing.Text
    ven_llave!VEM_DIRECCION = FrmVen.txtdireccion.Text
    ven_llave!VEM_TELE_CASA = FrmVen.txttelecasa.Text
    ven_llave!VEM_TELE_CELU = FrmVen.txttelecelu.Text
    ven_llave!VEM_SERIE_R = val(FrmVen.remi.Text)
    ven_llave!VEM_FLAG_G = " "
    ven_llave!VEM_FLAG_B = " "
    ven_llave!VEM_FLAG_F = " "
    'datos para app movil
    ven_llave!vem_hosting = 0
    ven_llave!vem_activo = Me.ComActivo.ListIndex
    ven_llave!vem_movil = Me.comLogeo.ListIndex
    ven_llave!vem_price = Me.comPrecio.ListIndex
    ven_llave!vem_idperfil = Me.ComPerfil.ListIndex
    ven_llave!VEM_IDEMPRESA = Me.DatEmpresa.BoundText

    If Me.comLogeo.ListIndex = 1 Then
        ven_llave!vem_login = Me.txtUser.Text

        If Len(Trim(Me.txtPass.Text)) <> 0 Then
    
            Dim cEncr As New CSHA256
    
            ven_llave!vem_pass = cEncr.SHADD256(Trim(Me.txtPass.Text))

        End If

    Else
        ven_llave!vem_login = ""
        ven_llave!vem_pass = ""

    End If
    'DATOS DEL REPARTIDOR - INICIO
    ven_llave!vem_brevete = Trim(Me.txtBrevete.Text)
    ven_llave!vem_placa = Trim(Me.txtPlaca.Text)
    ven_llave!vem_capacidad_kg = Me.txtCapacidad.Text
    'DATOS DEL REPARTIDOR - FIN

    If cheguia.Value = 1 Then
        ven_llave!VEM_FLAG_G = "A"

    End If

    If cheboleta.Value = 1 Then
        ven_llave!VEM_FLAG_B = "A"

    End If

    If chefactura.Value = 1 Then
        ven_llave!VEM_FLAG_F = "A"

    End If

    ven_llave("VEM_TRNKEY") = val(Right(cmbtransporte.Text, 10))
    NAMETRA = IIf(cmbtransporte.Text = "", " ", cmbtransporte.Text)
    ven_llave("VEM_TRANSPORTISTA") = Mid(NAMETRA, 1, 50)
    ven_llave.Update
    SQ_OPER = 2
    PUB_CODCIA = LK_CODCIA
    PUB_CODVEN = val(FrmVen.Txt_key.Text)
    LEER_PAR_LLAVE

    If pac_llave.EOF Then
        pac_llave.AddNew
    Else
        pac_llave.Edit

    End If

    pac_llave!pac_codcia = LK_CODCIA
    pac_llave!pac_codven = PUB_CODVEN
    pac_llave!PAC_ARCHI_F = txtfac.Text
    pac_llave!PAC_ARCHI_B = txtbol.Text
    pac_llave!PAC_ARCHI_G = txtguia.Text
    pac_llave!PAC_ARCHI_GUIA = txtgr.Text
    pac_llave!PAC_ARCHI_NC = txtnc.Text
    pac_llave!PAC_ARCHI_ND = txtnd.Text
    pac_llave!PAC_FLAG_CIA = " "
    pac_llave.Update

End Sub

Public Sub MENSAJE_VEN(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub

Public Sub LLENA_VEN(ban As Integer)
Dim I As Integer
If ban = 0 Then
       If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
         Else
          Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
       End If
       PUB_CODVEN = val(Txt_key.Text)
       pu_codcia = LK_CODCIA
       PUB_CODCIA = LK_CODCIA
       SQ_OPER = 1
       LEER_VEN_LLAVE
End If

FrmVen.Txt_key.Text = Trim(Nulo_Valors(ven_llave!VEM_CODVEN))
FrmVen.txtnombre.Text = Trim(Nulo_Valors(ven_llave!VEM_NOMBRE))
FrmVen.serie_g.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_G))

FrmVen.serie_nc.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_N))
FrmVen.serie_nd.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_D))
FrmVen.numfac_nc.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_N_INI))
FrmVen.numfac_nd.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_D_INI))
FrmVen.numfac_nc_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_N_FIN))
FrmVen.numfac_nd_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_D_FIN))

FrmVen.numfac_g.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_G_INI))

FrmVen.Serie_b.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_B))
FrmVen.serie_p.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_P))
FrmVen.numfac_b.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_B_INI))
FrmVen.serie_f.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_F))
FrmVen.numfac_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_F_INI))
FrmVen.numfac_p.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_P_INI))
FrmVen.numfac_g_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_G_FIN))
FrmVen.numfac_b_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_B_FIN))
FrmVen.numfac_f_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_F_FIN))
FrmVen.numfac_p_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_P_FIN))
If Not IsNull(ven_llave!VEM_FECHA_ING) Then
  txtfechaing.Text = Format(Nulo_Valors(ven_llave!VEM_FECHA_ING), "dd/mm/yyyy")
End If
txtfechaing.Mask = "##/##/####"
FrmVen.txtdireccion.Text = Trim(Nulo_Valors(ven_llave!VEM_DIRECCION))
FrmVen.txttelecasa.Text = Trim(Nulo_Valors(ven_llave!VEM_TELE_CASA))
FrmVen.txttelecelu.Text = Trim(Nulo_Valors(ven_llave!VEM_TELE_CELU))
FrmVen.remi.Text = Nulo_Valor0(ven_llave!VEM_SERIE_R)
'PARTE MOVIL
FrmVen.ComActivo.ListIndex = IIf(ven_llave!vem_activo, 1, 0)
FrmVen.comLogeo.ListIndex = IIf(ven_llave!vem_movil, 1, 0)
FrmVen.txtUser.Text = ven_llave!vem_login
FrmVen.comPrecio.ListIndex = IIf(ven_llave!vem_price, 1, 0)
FrmVen.ComPerfil.ListIndex = IIf(Trim(Nulo_Valors(ven_llave!vem_idperfil)) = "", 0, ven_llave!vem_idperfil)
FrmVen.DatEmpresa.BoundText = ven_llave!VEM_IDEMPRESA
'PARTE MOVIL FIN
'DATOS DE REPARTIDOR - INICIO
Me.txtBrevete.Text = Nulo_Valors(ven_llave!vem_brevete)
Me.txtPlaca.Text = Nulo_Valors(ven_llave!vem_placa)
Me.txtCapacidad.Text = Nulo_Valor0(ven_llave!vem_capacidad_kg)
'DATOS DE REPARTIDOR - FIN
FindInCmb Nulo_Valor0(ven_llave!VEM_TRNKEY)
cheguia.Value = 0
cheboleta.Value = 0
chefactura.Value = 0
Check1.Value = 0
chenc.Value = 0
chend.Value = 0
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_G)) = "A" Then
  cheguia.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_B)) = "A" Then
  cheboleta.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_F)) = "A" Then
  chefactura.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_P)) = "A" Then
  Check1.Value = 1
End If

If UCase(Nulo_Valors(ven_llave!VEM_FLAG_N)) = "A" Then
  chenc.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_D)) = "A" Then
  chend.Value = 1
End If
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
PUB_CODVEN = val(ven_llave!VEM_CODVEN)
LEER_PAR_LLAVE
If Not pac_llave.EOF Then
 txtfac.Text = Trim(pac_llave!PAC_ARCHI_F)
 txtbol.Text = Trim(pac_llave!PAC_ARCHI_B)
 txtguia.Text = Trim(pac_llave!PAC_ARCHI_G)
 txtgr.Text = Trim(pac_llave!PAC_ARCHI_GUIA)
 txtnc.Text = Trim(pac_llave!PAC_ARCHI_NC)
 txtnd.Text = Trim(pac_llave!PAC_ARCHI_ND)
End If


End Sub
Public Sub LIMPIA_VEN()
Txt_key.Text = ""
txtnombre.Text = ""
serie_g.Text = ""
numfac_g.Text = ""
Serie_b.Text = ""
numfac_b.Text = ""
serie_f.Text = ""
serie_nc.Text = ""
serie_nd.Text = ""

numfac_f.Text = ""
numfac_nc.Text = ""
numfac_nd.Text = ""

numfac_g_f.Text = ""
numfac_b_f.Text = ""
numfac_f_f.Text = ""
numfac_nc_f.Text = ""
numfac_nd_f.Text = ""

cheguia.Value = 0
cheboleta.Value = 0
chefactura.Value = 0
chenc.Value = 0
chend.Value = 0

Check1.Value = 0
serie_p.Text = ""
numfac_p.Text = ""
numfac_p_f.Text = ""
remi.Text = ""

txtfechaing.Text = "00/00/0000"
FrmVen.txtdireccion.Text = ""
FrmVen.txttelecasa.Text = ""
FrmVen.txttelecelu.Text = ""


txtfac.Text = ""
txtbol.Text = ""
txtguia.Text = ""
txtgr.Text = ""
txtnc.Text = ""
txtnd.Text = ""

Me.txtUser.Text = ""
Me.txtPass = ""
Me.txtBrevete.Text = ""
Me.txtPlaca.Text = ""
Me.txtCapacidad.Text = 0

cmbtransporte.ListIndex = -1
Me.ComPerfil.ListIndex = 0
End Sub

Private Sub cheboleta_Click()
If Serie_b.Enabled Then
 Serie_b.SetFocus
End If
End Sub

Private Sub chefactura_Click()
If serie_f.Enabled Then
 serie_f.SetFocus
End If
End Sub

Private Sub cheguia_Click()
If serie_g.Enabled Then
 serie_g.SetFocus
End If
End Sub

Private Sub cmdagregar_Click()

    'On Error GoTo ESCAPA
    If Left(cmdAgregar.Caption, 2) = "&A" Then
        cmdAgregar.Caption = "&Grabar"
        cmdCancelar.Enabled = True
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
        LIMPIA_VEN
        DESBLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
        DESBLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
        DESBLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
        remi.Enabled = True
        txtfechaing.Enabled = True
        FrmVen.Txt_key = GENERA_VEN
        txtfechaing.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
        Me.comLogeo.ListIndex = 1
        Me.ComActivo.ListIndex = 1
        Me.txtUser.Enabled = True
        Me.txtUser.Text = ""
        Me.txtPass.Text = ""
        FrmVen.txtnombre.SetFocus
        'AGREGAMOS EN BLANCO
    Else

        'VALIDA SI EL USUARIO EXISTE EN CLOUD\
        MousePointer = vbHourglass

        If Len(Trim(Me.txtUser.Text)) <> 0 Then
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[dbo].[USP_USUARIO_VALIDA_REGISTER]"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USER", adVarChar, adParamInput, 20, Me.txtUser.Text)

            Dim orsValida As ADODB.Recordset

            Set orsValida = oCmdEjec.Execute
        
            If Not orsValida.EOF Then
                MousePointer = vbDefault

                If orsValida!Dato = 1 Then
                    MsgBox "Usuario ya se encuentra registrado."
                    Exit Sub

                End If

            End If

        End If

        If FrmVen.txtnombre.Text = "" Or Len(FrmVen.txtnombre.Text) = 0 Then
            MsgBox "Ingrese Nombre de Vendedor ..!!!", 48, Pub_Titulo
            Azul txtnombre, txtnombre
            Exit Sub

        End If

        If Me.comLogeo.ListIndex = 1 And Len(Trim(Me.txtUser.Text)) = 0 Then
            MsgBox "Debe ingresar el Usuario del Vendedor.", vbCritical, Pub_Titulo
            Me.txtUser.SetFocus
            Exit Sub

        End If

        If Me.comLogeo.ListIndex = 1 And Len(Trim(Me.txtPass.Text)) = 0 Then
            MsgBox "Debe ingresar el Pass del Vendedor.", vbCritical, Pub_Titulo
            Me.txtPass.SetFocus
            Exit Sub

        End If

        If Me.ComPerfil.ListIndex = 0 Then
            MsgBox "Debe elegir el perfil del Usuario.", vbCritical, Pub_Titulo
            Me.ComPerfil.SetFocus
            Exit Sub

        End If

        WSFECHA = ES_FECHAS(txtfechaing)

        If WSFECHA = "1" Then
            MsgBox " Fecha Invalidad ...", 48, Pub_Titulo
            Azul2 txtfechaing, txtfechaing
            Exit Sub

        End If
        
        'VALIDACION PARA TRANSPORTISTA - INICIO
        If Me.ComPerfil.ListIndex = 3 And Len(Trim(Me.txtBrevete.Text)) = 0 Then
            Me.SSTab1.tab = 2
            MsgBox "Debe ingresar el Brevete del Repartidor", vbCritical, Pub_Titulo
            Me.txtBrevete.SetFocus
            Exit Sub
        End If
        
        If Me.ComPerfil.ListIndex = 3 And Len(Trim(Me.txtPlaca.Text)) = 0 Then
        Me.SSTab1.tab = 2
            MsgBox "Debe ingresar la Placa del vehiculo.", vbCritical, Pub_Titulo
            Me.txtPlaca.SetFocus
            Exit Sub
        End If
        
        If Me.ComPerfil.ListIndex = 3 And Len(Trim(Me.txtCapacidad.Text)) = 0 Then
        Me.SSTab1.tab = 2
            MsgBox "Debe ingresar la capacidad del vehiculo.", vbCritical, Pub_Titulo
            Me.txtCapacidad.SetFocus
            Exit Sub
        End If
        
        If Me.ComPerfil.ListIndex = 3 And val(Me.txtCapacidad.Text) <= 0 Then
        Me.SSTab1.tab = 2
            MsgBox "Capacidad ingresada incorrecta.", vbInformation, Pub_Titulo
            Me.txtCapacidad.SetFocus
            Exit Sub
        End If
        'VALIDACION PARA TRANSPORTISTA - FIN

        txtfechaing.Text = Format(WSFECHA, "dd/mm/yyyy")
        '"SI GRABA.."
        SQ_OPER = 1
        PUB_CODVEN = val(FrmVen.Txt_key.Text)
        pu_codcia = LK_CODCIA
        LEER_VEN_LLAVE

        If Not ven_llave.EOF Then
            MsgBox "Registro ,  EXISTE ... ", 48, Pub_Titulo
            Azul FrmVen.Txt_key, Txt_key
            Exit Sub

        End If

        Screen.MousePointer = 11
        GRABAR_VEN
        MENSAJE_VEN "Bancos , AGREGADO... "
        cmdAgregar.Caption = "&Agregar"
        cmdEliminar.Enabled = True
        cmdModificar.Enabled = True
        LIMPIA_VEN
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
        BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
        remi.Enabled = False
        txtfechaing.Enabled = False
        Txt_key.Locked = False
        Txt_key.SetFocus
        Screen.MousePointer = 0

    End If
   
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    
End If

End Sub

Private Sub cmdCancelar_Click()
If Left(cmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
    LIMPIA_VEN
    Txt_key.Locked = False
    MENSAJE_VEN "Proceso Cancelado... !!!    "
    Txt_key.Enabled = True
    Txt_key.SetFocus
     Exit Sub
End If
     Screen.MousePointer = 11
     If Left(cmdModificar.Caption, 2) = "&G" Then
        cmdModificar.Caption = "&Modificar"
        LLENA_VEN 1
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
        BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
        remi.Enabled = False
        txtfechaing.Enabled = False
        
        Txt_key.Locked = True
     Else
        cmdAgregar.Caption = "&Agregar"
        LIMPIA_VEN
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
        BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
        remi.Enabled = False
        txtfechaing.Enabled = False
        Txt_key.Locked = False
     End If
     cmdCerrar.Caption = "&Cerrar"
     cmdCancelar.Enabled = True
     cmdAgregar.Enabled = True
     cmdModificar.Enabled = True
     cmdEliminar.Enabled = True
     Txt_key.Enabled = True
     MENSAJE_VEN "Proceso Cancelado... !!!    "
     Txt_key.SetFocus
     Screen.MousePointer = 0

End Sub

Private Sub cmdCerrar_Click()
ws_conta = 0
Unload FrmVen

End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    FrmVen.Txt_key.SetFocus
End If

End Sub

Private Sub cmdEliminar_Click()
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset

If Len(Txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_VEN "NO a seleccionado NADA ... !"
   Exit Sub
End If
  pub_cadena = "SELECT FAR_CODVEN FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODVEN = ? "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01(0) = 0
  PS_REP01(1) = 0
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ven_llave!VEM_CODVEN
  llave_rep01.Requery
  If Not llave_rep01.EOF Then
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  Vendedor  TIENE H I S T O R I A.. ", 48, Pub_Titulo
     Exit Sub
  End If
  
  pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    ven_llave.Delete
    Txt_key.Text = ""
    Txt_key.Locked = False
    LIMPIA_VEN
    MENSAJE_VEN "Registro   ELIMINADO ... "
    Screen.MousePointer = 0
   Exit Sub
  End If
  Screen.MousePointer = 0
End Sub

Private Sub CmdModificar_Click()
If Len(Txt_key) = 0 Then
   MENSAJE_VEN "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(cmdModificar.Caption, 2) = "&M" Then
    cmdModificar.Caption = "&Grabar"
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    Txt_key.Locked = True
    DESBLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
    DESBLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
    DESBLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
    remi.Enabled = True
    txtfechaing.Enabled = True
    Me.txtUser.Enabled = False
    txtnombre.SetFocus
Else
    '*Grabar las modificaciones
    If txtnombre.Text = "" Or Len(txtnombre.Text) = 0 Then
         MsgBox " Nombre Invalido ....", 48, Pub_Titulo
         Exit Sub
    End If
    WSFECHA = ES_FECHAS(txtfechaing)
    If WSFECHA = "1" Then
      MsgBox " Fecha Invalidad ...", 48, Pub_Titulo
      Azul2 txtfechaing, txtfechaing
      Exit Sub
    End If
    If Me.ComPerfil.ListIndex = 0 Then
        MsgBox "Debe elegir el perfil del usuario.", vbCritical, Pub_Titulo
        Me.ComPerfil.SetFocus
        Exit Sub
    End If
    'VALIDACIONES DEL REPARTIDOR - INICIO
    If Me.ComPerfil.ListIndex = 3 And Len(Trim(Me.txtBrevete.Text)) = 0 Then
            Me.SSTab1.tab = 2
            MsgBox "Debe ingresar el Brevete del Repartidor", vbCritical, Pub_Titulo
            Me.txtBrevete.SetFocus
            Exit Sub
        End If
        
        If Me.ComPerfil.ListIndex = 3 And Len(Trim(Me.txtPlaca.Text)) = 0 Then
        Me.SSTab1.tab = 2
            MsgBox "Debe ingresar la Placa del vehiculo.", vbCritical, Pub_Titulo
            Me.txtPlaca.SetFocus
            Exit Sub
        End If
        
        If Me.ComPerfil.ListIndex = 3 And Len(Trim(Me.txtCapacidad.Text)) = 0 Then
        Me.SSTab1.tab = 2
            MsgBox "Debe ingresar la capacidad del vehiculo.", vbCritical, Pub_Titulo
            Me.txtCapacidad.SetFocus
            Exit Sub
        End If
        
        If Me.ComPerfil.ListIndex = 3 And val(Me.txtCapacidad.Text) <= 0 Then
        Me.SSTab1.tab = 2
            MsgBox "Capacidad ingresada incorrecta.", vbInformation, Pub_Titulo
            Me.txtCapacidad.SetFocus
            Exit Sub
        End If
    'VALIDACIONES DEL REPARTIDOR - FIN
    txtfechaing.Text = Format(WSFECHA, "dd/mm/yyyy")
     Screen.MousePointer = 11
     GRABAR_VEN
     MENSAJE_VEN "Registro , MODIFICADO... "
     cmdModificar.Caption = "&Modificar"
     cmdCancelar.Enabled = True
     cmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     Txt_key.Locked = True
     BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
     BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
     BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
     remi.Enabled = False
     txtfechaing.Enabled = False
     Screen.MousePointer = 0
End If

End Sub

Private Sub ComPerfil_Click()
If Me.ComPerfil.ListIndex = 3 Then
    Me.SSTab1.TabVisible(2) = True
Else
    Me.SSTab1.TabVisible(2) = False
End If
End Sub

Private Sub Form_Load()
Unload FORMGEN
If LK_CODCIA = "04" Then
'  FrmVen.Caption = "&Chofer / Solic."
'  F1.Caption = "&Chofer / Solic."
Else
'  FrmVen.Caption = "&Vendedor"
 ' F1.Caption = "Vendedor"
End If

loc_key = 0
LIMPIA_VEN
BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
remi.Enabled = False
txtfechaing.Enabled = False
Txt_key.Enabled = True
F2.Visible = False
If LK_FLAG_FACTURACION = "V" Then
  F2.Visible = True
End If

LlenaTransporte
LlenaEmpresa
pub_cadena = "SELECT * FROM TRANSPORTE WHERE TRN_KEY = ? ORDER BY TRN_NOMBRE"
Set PS_TRAONE = CN.CreateQuery("", pub_cadena)
PS_TRAONE(0) = 0
Set TRANSPORTEONE = PS_TRAONE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

Me.SSTab1.TabVisible(2) = False
End Sub

Private Sub LlenaEmpresa()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_EMPRESA_LIST]"
Dim ORSdatos  As ADODB.Recordset
Set ORSdatos = oCmdEjec.Execute
Set Me.DatEmpresa.RowSource = ORSdatos
Me.DatEmpresa.ListField = ORSdatos(1).Name
Me.DatEmpresa.BoundColumn = ORSdatos(0).Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
ws_conta = 0
End Sub

Public Sub BLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10, Optional o11, Optional o12)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = False
' o1.BackColor = QBColor(7)
End If
If Not IsMissing(o2) Then
 o2.Enabled = False
 'o2.BackColor = QBColor(7)
End If
If Not IsMissing(o3) Then
 o3.Enabled = False
 'o3.BackColor = QBColor(7)
End If
If Not IsMissing(o4) Then
 o4.Enabled = False
 'o4.BackColor = QBColor(7)
End If
If Not IsMissing(o5) Then
 o5.Enabled = False
 'o5.BackColor = QBColor(7)
End If
If Not IsMissing(o6) Then
 o6.Enabled = False
 'o6.BackColor = QBColor(7)
End If
If Not IsMissing(o7) Then
 o7.Enabled = False
 'o7.BackColor = QBColor(7)
End If
If Not IsMissing(o8) Then
 o8.Enabled = False
 'o8.BackColor = QBColor(7)
End If
If Not IsMissing(o9) Then
 o9.Enabled = False
 'o9.BackColor = QBColor(7)
End If
If Not IsMissing(o10) Then
 o10.Enabled = False
 'o10.BackColor = QBColor(7)
End If
If Not IsMissing(o11) Then
 o11.Enabled = False
 'o11.BackColor = QBColor(7)
End If
If Not IsMissing(o12) Then
 o12.Enabled = False
 'o12.BackColor = QBColor(7)
End If

End Sub
Public Sub DESBLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10, Optional o11, Optional o12)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = True
' o1.BackColor = QBColor(15)
End If
If Not IsMissing(o2) Then
 o2.Enabled = True
' o2.BackColor = QBColor(15)
End If
If Not IsMissing(o3) Then
 o3.Enabled = True
' o3.BackColor = QBColor(15)
End If
If Not IsMissing(o4) Then
 o4.Enabled = True
' o4.BackColor = QBColor(15)
End If
If Not IsMissing(o5) Then
 o5.Enabled = True
' o5.BackColor = QBColor(15)
End If
If Not IsMissing(o6) Then
 o6.Enabled = True
' o6.BackColor = QBColor(15)
End If
If Not IsMissing(o7) Then
 o7.Enabled = True
' o7.BackColor = QBColor(15)
End If
If Not IsMissing(o8) Then
 o8.Enabled = True
' o8.BackColor = QBColor(15)
End If
If Not IsMissing(o9) Then
 o9.Enabled = True
' o9.BackColor = QBColor(15)
End If
If Not IsMissing(o10) Then
 o10.Enabled = True
' o10.BackColor = QBColor(15)
End If
If Not IsMissing(o11) Then
 o11.Enabled = True
' o11.BackColor = QBColor(15)
End If
If Not IsMissing(o12) Then
 o12.Enabled = True
' o12.BackColor = QBColor(15)
End If

End Sub




Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
txt_key_KeyPress 13
End Sub

Private Sub numfac_b_f_GotFocus()
Azul numfac_b_f, numfac_b_f
End Sub

Private Sub numfac_b_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 serie_f.SetFocus
End If

End Sub

Private Sub numfac_b_GotFocus()
Azul numfac_b, numfac_b
End Sub

Private Sub numfac_b_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_b_f.SetFocus
End If

End Sub

Private Sub numfac_f_f_GotFocus()
Azul numfac_f_f, numfac_f_f
End Sub

Private Sub numfac_f_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
  Exit Sub
End If
If cmdModificar.Enabled Then
   cmdModificar.SetFocus
Else
   cmdAgregar.SetFocus
End If

End Sub

Private Sub numfac_f_GotFocus()
Azul numfac_f, numfac_f
End Sub

Private Sub numfac_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  numfac_f_f.SetFocus
End If
End Sub

Private Sub numfac_g_f_GotFocus()
Azul numfac_g_f, numfac_g_f
End Sub

Private Sub numfac_g_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 Serie_b.SetFocus
End If

End Sub

Private Sub numfac_g_GotFocus()
Azul numfac_g, numfac_g
End Sub

Private Sub numfac_g_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_g_f.SetFocus
End If

End Sub

Private Sub remi_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub Serie_b_GotFocus()
Azul Serie_b, Serie_b
End Sub

Private Sub Serie_b_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_b.SetFocus
End If

End Sub

Private Sub serie_f_GotFocus()
Azul serie_f, serie_f
End Sub

Private Sub serie_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_f.SetFocus
End If

End Sub

Private Sub serie_g_GotFocus()
Azul serie_g, serie_g
End Sub

Private Sub serie_g_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_g.SetFocus
End If
End Sub



Private Sub txt_key_GotFocus()
 Azul Txt_key, Txt_key
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And Txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

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
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  Txt_key.SelStart = Len(Txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)

Dim valor As String
Dim tf As Integer
Dim I
Dim itmFound As ListItem
If KeyAscii = 27 And Trim(txtnombre.Text) = "" Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
pu_codclie = val(Txt_key.Text)
If Len(Txt_key.Text) = 0 Or Txt_key.Locked Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(Txt_key.Text) = True Then
   loc_key = 0
   On Error GoTo mucho
   PUB_CODVEN = val(Txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   ListView1.Visible = False
   cmdCancelar.Enabled = True
   LLENA_VEN 0
   Txt_key.Locked = True
   cmdModificar.SetFocus
   Screen.MousePointer = 0
   Exit Sub
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(Txt_key.Text)) = Left(valor, Len(Trim(Txt_key.Text))) Then
   Else
      Exit Sub
   End If
   ListView1.Visible = False
   cmdCancelar.Enabled = True
   LLENA_VEN 0
    Txt_key.Locked = True
   cmdCancelar.Enabled = True
    cmdModificar.SetFocus
End If
dale:
mucho:
ListView1.Visible = False
fin:
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(Txt_key.Text) = 0 Or Txt_key.Locked = True Or IsNumeric(Txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key.Text) = 1 Then
    var = Asc(Txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    archi = "SELECT * FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If ListView1.Visible = False Then
        loc_key = 0
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(Txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If
End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = True 'Not LblMensaje.Visible
 If CU > 8 Then
   CU = 0
   PARPADEA.Enabled = False
   LblMensaje.Visible = False
 End If
End Sub

Private Sub txtdireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul2 txtfechaing, txtfechaing
End If
End Sub

Private Sub txtfechaing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txttelecasa, txttelecasa
End If
End Sub

Private Sub txtnombre_GotFocus()
Azul txtnombre, txtnombre
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txtdireccion, txtdireccion
End If
End Sub

Public Function ES_FECHAS(CAMPOFECHA As MaskEdBox) As String
Dim wfecha As String
ES_FECHAS = "0"
If CAMPOFECHA = "00/00/0000" Then
 Exit Function
End If
If Right(CAMPOFECHA.Text, 2) = "__" Then
  wfecha = Left(CAMPOFECHA.Text, 8)
Else
  wfecha = Trim(CAMPOFECHA.Text)
End If
If Not IsDate(wfecha) Then
  ES_FECHAS = "1"
  Exit Function
End If
ES_FECHAS = wfecha
End Function

Private Sub txttelecasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txttelecelu, txttelecelu
End If
End Sub

Private Sub txttelecelu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If F2.Visible Then
    Azul serie_g, serie_g
  Else
   If cmdModificar.Enabled Then
    cmdModificar.SetFocus
   Else
    cmdAgregar.SetFocus
   End If
  End If

End If
End Sub

Private Sub LlenaTransporte()
Dim PS_TRA As rdoQuery
Dim TRANSPORTE As rdoResultset
Dim SQL As String
SQL = "SELECT * FROM TRANSPORTE ORDER BY TRN_NOMBRE"
Set PS_TRA = CN.CreateQuery("", SQL)
Set TRANSPORTE = PS_TRA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
TRANSPORTE.Requery
cmbtransporte.Clear
Do Until TRANSPORTE.EOF
    cmbtransporte.AddItem Trim(TRANSPORTE!TRN_NOMBRE) & String(80, " ") & TRANSPORTE!TRN_KEY
    TRANSPORTE.MoveNext
Loop

End Sub
Private Function FindInCmb(ByVal s_transporte As String) As Boolean
Dim I As Long
Dim aux_f As String

    cmbtransporte.ListIndex = -1
    For I = 0 To cmbtransporte.ListCount - 1
     aux_f = cmbtransporte.List(I)
     aux_f = Trim$(Right$(aux_f, 10))
     If Trim(aux_f) = Trim(s_transporte) Then
      cmbtransporte.ListIndex = I
      Exit For
     End If
    Next
   
End Function
