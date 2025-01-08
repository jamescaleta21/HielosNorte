VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepartos_Pedido 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalles de Pedido:"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10350
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
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10095
      Begin VB.Label lblObs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2760
         TabIndex        =   7
         Top             =   1200
         Width           =   3945
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2760
         TabIndex        =   6
         Top             =   720
         Width           =   3945
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   3945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
         Height          =   195
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   675
      End
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5953
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
End
Attribute VB_Name = "frmRepartos_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gIDpedido As Double

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
ConfigurarLV
cargarPedido
Me.Caption = "Detalle de Pedido: " & gIDpedido
End Sub


Private Sub ConfigurarLV()
With Me.lvDetalle
.ColumnHeaders.Add , , "Cant"
.ColumnHeaders.Add , , "Producto", 3500
.ColumnHeaders.Add , , "Precio"
.ColumnHeaders.Add , , "Importe"
.FullRowSelect = True
.HideColumnHeaders = False
.View = lvwReport
End With
End Sub

Private Sub cargarPedido()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "USP_PEDIDO_FILL"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPEDIDO", adBigInt, adParamInput, , gIDpedido)
            
    Dim orsDatos As ADODB.Recordset

    Set orsDatos = oCmdEjec.Execute
    
    Me.lblCliente.Caption = orsDatos!cliente
    Me.lblDireccion.Caption = orsDatos!dir
    Me.lblObs.Caption = orsDatos!obs
    
    Dim ORSt As ADODB.Recordset
    Set ORSt = orsDatos.NextRecordset
    
    Do While Not ORSt.EOF
     Set itemX = Me.lvDetalle.ListItems.Add(, , ORSt!cant)
        itemX.Tag = ORSt!IDEPRODUCTO
        itemX.SubItems(1) = ORSt!PRODUCTO
        itemX.SubItems(2) = ORSt!PRE
        itemX.SubItems(3) = ORSt!imp
        ORSt.MoveNext
    Loop

End Sub

