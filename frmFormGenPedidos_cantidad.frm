VERSION 5.00
Begin VB.Form frmFormGenPedidos_cantidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   720
      Left            =   2880
      Picture         =   "frmFormGenPedidos_cantidad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   720
      Left            =   2880
      Picture         =   "frmFormGenPedidos_cantidad.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtPrecio 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtCantidad 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   765
   End
End
Attribute VB_Name = "frmFormGenPedidos_cantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gacepta As Boolean
Public gCantidad As Integer
Public gPrecio As Double

Private Sub cmdAceptar_Click()
If Len(Trim(Me.txtCantidad.Text)) = 0 Then
    MsgBox "Ingrese Cantidad.", vbCritical, Pub_Titulo
    Me.txtCantidad.SetFocus
    Exit Sub
End If
If Len(Trim(Me.txtPrecio.Text)) = 0 Then
    MsgBox "Ingrese Precio.", vbCritical, Pub_Titulo
    Me.txtPrecio.SetFocus
    Exit Sub
End If
If IsNumeric(Me.txtCantidad.Text) = False Then
    MsgBox "Debe ingresar cantidad correcta", vbCritical, Pub_Titulo
    Me.txtCantidad.SetFocus
    Me.txtCantidad.SelStart = 0
    Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
    Exit Sub
End If
If IsNumeric(Me.txtPrecio.Text) = False Then
    MsgBox "Debe ingresar Precio correcto", vbCritical, Pub_Titulo
    Me.txtPrecio.SetFocus
    Me.txtPrecio.SelStart = 0
    Me.txtPrecio.SelLength = Len(Me.txtPrecio.Text)
    Exit Sub
End If
gacepta = True
gCantidad = Me.txtCantidad.Text
gPrecio = Me.txtPrecio.Text
cmdCancelar_Click
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtPrecio.SetFocus
    Me.txtPrecio.SelStart = 0
    Me.txtPrecio.SelLength = Len(Me.txtPrecio.Text)
End If
End Sub
