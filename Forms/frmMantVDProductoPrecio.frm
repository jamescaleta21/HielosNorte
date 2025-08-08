VERSION 5.00
Begin VB.Form frmMantVDProductoPrecio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizacion de Datos"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5955
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
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   1200
         Left            =   240
         Picture         =   "frmMantVDProductoPrecio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   1200
         Left            =   240
         Picture         =   "frmMantVDProductoPrecio.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox ComActivo 
         Height          =   360
         ItemData        =   "frmMantVDProductoPrecio.frx":1994
         Left            =   1800
         List            =   "frmMantVDProductoPrecio.frx":199E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtAddPrecio 
         Height          =   360
         Left            =   1800
         TabIndex        =   0
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblCategoria 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   240
         Left            =   840
         TabIndex        =   7
         Top             =   1980
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio:"
         Height          =   240
         Left            =   840
         TabIndex        =   5
         Top             =   1260
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmMantVDProductoPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gIDcategoria As Integer
Public gIDempresa As Integer

Private Sub cmdAceptar_Click()
With frmMantVDProducto.lvPrecios.SelectedItem
    .Text = Me.lblCategoria.Caption
    .SubItems(1) = Me.lblCategoria.Tag
    .SubItems(2) = Me.txtAddPrecio.Text
    .SubItems(3) = IIf(Me.ComActivo.ListIndex = 0, "NO", "SI")
End With

Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub ComActivo_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.cmdAceptar
End Sub

Private Sub datAddCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub datAddCategoria_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.txtAddPrecio
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
gacepta = False
Me.lblCategoria.Tag = gIDcategoria
End Sub

Private Sub mebFin_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.ComActivo
End Sub

Private Sub mebIni_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.mebFin
End Sub

Private Sub txtAddPrecio_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.ComActivo
End Sub
