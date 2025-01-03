VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepartos_Pedido 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalles de Pedido:"
   ClientHeight    =   4950
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
   ScaleHeight     =   4950
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8070
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
