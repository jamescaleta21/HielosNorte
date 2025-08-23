VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormGenPedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11655
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
   ScaleHeight     =   7755
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   720
      Left            =   10320
      Picture         =   "frmFormGenPedidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   720
      Left            =   9120
      Picture         =   "frmFormGenPedidos.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   11415
      Begin MSComctlLib.ListView lvSearch 
         Height          =   2055
         Left            =   1200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3625
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
      Begin VB.TextBox txtPrecio 
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtProducto 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   7335
      End
      Begin VB.CommandButton cmdDel 
         Enabled         =   0   'False
         Height          =   480
         Left            =   10200
         Picture         =   "frmFormGenPedidos.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   480
         Left            =   10200
         Picture         =   "frmFormGenPedidos.frx":163E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   990
      End
      Begin MSComctlLib.ListView lvDetalle 
         Height          =   2775
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4895
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
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   8880
         TabIndex        =   22
         Top             =   3960
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   8280
         TabIndex        =   21
         Top             =   4035
         Width           =   495
      End
      Begin VB.Label lblidprod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   9120
         TabIndex        =   15
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblImporte 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6840
         TabIndex        =   14
         Top             =   705
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         Height          =   195
         Left            =   5160
         TabIndex        =   13
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio:"
         Height          =   195
         Left            =   2640
         TabIndex        =   12
         Top             =   765
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   765
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
         Height          =   195
         Left            =   375
         TabIndex        =   10
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11415
      Begin VB.Label lblObservacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   1560
         TabIndex        =   29
         Top             =   1320
         Width           =   7275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblIDvendedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         Height          =   195
         Left            =   8880
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblIDcliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         Height          =   195
         Left            =   7440
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblIDpedido 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8880
         TabIndex        =   24
         Top             =   600
         Width           =   2220
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   615
         TabIndex        =   20
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4560
         TabIndex        =   19
         Top             =   420
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   840
         Width           =   2835
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5400
         TabIndex        =   16
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   2835
      End
   End
End
Attribute VB_Name = "frmFormGenPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gIDpedido As Double
Private loc_key  As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes
Public vGraba As Boolean 'variable para validar si pulsó el boton grabar
Public vTotal As Double 'variable para almacenar el total del pedido para enviarlo al listviewpedidos del formgen

Private Sub cmdAdd_Click()

    If Len(Trim(Me.txtProducto.Text)) = 0 Then
        MsgBox "Debe elegir el producto antes de agregar.", vbInformation, Pub_Titulo
        Me.txtProducto.SetFocus
        Exit Sub

    End If

    If Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Debe ingresar la cantidad antes de agregar.", vbInformation, Pub_Titulo
        Me.txtCantidad.SetFocus
        Exit Sub

    End If

    If Len(Trim(Me.txtPrecio.Text)) = 0 Then
        MsgBox "Debe ingresar el precio antes de agregar.", vbInformation, Pub_Titulo
        Me.txtPrecio.SetFocus
        Exit Sub

    End If

    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "Cantidad ingresada incorrecta.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
        Exit Sub

    End If

    If Not IsNumeric(Me.txtPrecio.Text) Then
        MsgBox "Precio ingresado incorrecto.", vbCritical, Pub_Titulo
        Me.txtPrecio.SetFocus
        Me.txtPrecio.SelStart = 0
        Me.txtPrecio.SelLength = Len(Me.txtCantidad.Text)
        Exit Sub

    End If

    Dim itemx As Object
    
    If Me.lvDetalle.ListItems.count = 0 Then
        Set itemx = Me.lvDetalle.ListItems.Add(, , Me.txtCantidad.Text)
        itemx.Tag = Me.lblidprod.Caption
        itemx.SubItems(1) = Me.txtProducto.Text
        itemx.SubItems(2) = Me.txtPrecio.Text
        itemx.SubItems(3) = Me.lblImporte.Caption
    Else

        Dim vEncontrado As Boolean

        vEncontrado = False

        For Each itemx In Me.lvDetalle.ListItems

            If itemx.Tag = Me.txtProducto.Tag Then
                vEncontrado = True
                Exit For

            End If

        Next

        If vEncontrado = False Then
            Set itemx = Me.lvDetalle.ListItems.Add(, , Me.txtCantidad.Text)
            itemx.Tag = Me.lblidprod.Caption
            itemx.SubItems(1) = Me.txtProducto.Text
            itemx.SubItems(2) = Me.txtPrecio.Text
            itemx.SubItems(3) = Me.lblImporte.Caption
        Else
            MsgBox "Producto ya agregado", vbCritical, Pub_Titulo

        End If

    End If
    
    Me.txtCantidad.Text = ""
    Me.txtPrecio.Text = ""
    Me.txtProducto.Tag = ""
    Me.lblImporte.Caption = "0.00"
    Me.txtProducto.Text = ""
    Me.txtProducto.SetFocus
CalculoTotal
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdDel_Click()
Me.lvDetalle.ListItems.Remove Me.lvDetalle.SelectedItem.Index
Me.cmdDel.Enabled = False
CalculoTotal
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo cGraba
Pub_ConnAdo.BeginTrans
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_BACKUP]"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPEDIDO", adBigInt, adParamInput, , Me.lblIDpedido.Caption)
oCmdEjec.Execute

LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_UPDATE]"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPEDIDO", adBigInt, adParamInput, , Me.lblIDpedido.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, 2, Me.lblIDcliente.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTotal.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.lblIDvendedor.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 20, LK_CODUSU)
oCmdEjec.Execute

LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_ELIMINADETALLE]"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPEDIDO", adBigInt, adParamInput, , Me.lblIDpedido.Caption)
oCmdEjec.Execute

LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_REGISTRADETALLE]"
Dim i As Integer
For i = 1 To Me.lvDetalle.ListItems.count
    LimpiaParametros oCmdEjec
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPEDIDO", adBigInt, adParamInput, , Me.lblIDpedido.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lvDetalle.ListItems(i).Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SECUENCIA", adBigInt, adParamInput, , i)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adDouble, adParamInput, , Me.lvDetalle.ListItems(i).Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRECIO", adDouble, adParamInput, , Me.lvDetalle.ListItems(i).SubItems(2))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IMPORTE", adDouble, adParamInput, , Me.lvDetalle.ListItems(i).SubItems(3))
    oCmdEjec.Execute
Next

Pub_ConnAdo.CommitTrans
vGraba = True
vTotal = Me.lblTotal.Caption
MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
cmdCancelar_Click
Exit Sub
cGraba:
Pub_ConnAdo.RollbackTrans
MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
ConfiguraLV
vBuscar = False
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_FILL]"

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPEDIDO", adBigInt, adParamInput, , gIDpedido)

Dim ORSd As ADODB.Recordset
Set ORSd = oCmdEjec.Execute

If Not ORSd.EOF Then
    Me.lblCliente.Caption = ORSd!cliente
    Me.lblIDcliente.Caption = ORSd!idcliente
    Me.lblIDvendedor.Caption = ORSd!IDVEN
    Me.lblVendedor.Caption = ORSd!vendedor
    Me.lblIDpedido.Caption = ORSd!ide
    Me.lblFecha.Caption = ORSd!fecha
    Me.lblTotal.Caption = ORSd!Total
    Me.lblObservacion.Caption = ORSd!obs
End If

Dim ORSt As ADODB.Recordset
Set ORSt = ORSd.NextRecordset

Dim itemx As Object
Do While Not ORSt.EOF
    Set itemx = Me.lvDetalle.ListItems.Add(, , ORSt!cant)
    itemx.Tag = ORSt!IDEPRODUCTO
    itemx.SubItems(1) = ORSt!producto
    itemx.SubItems(2) = ORSt!PRE
    itemx.SubItems(3) = ORSt!imp
    ORSt.MoveNext
Loop
vBuscar = True
End Sub

Private Sub ConfiguraLV()
With Me.lvDetalle
     .ColumnHeaders.Add , , "Cantidad", 1000
    .ColumnHeaders.Add , , "Artículo", 6000
    .ColumnHeaders.Add , , "Precio", 1000
    .ColumnHeaders.Add , , "Importe", 1000
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = False
    .View = lvwReport
    .HideSelection = False
End With
With Me.lvSearch
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 800
    .ColumnHeaders.Add , , "Articulo", 4000
    .ColumnHeaders.Add , , "Precio", 800
    .MultiSelect = False
End With
End Sub







Private Sub lvDetalle_DblClick()
frmFormGenPedidos_cantidad.txtCantidad = Me.lvDetalle.SelectedItem.Text
frmFormGenPedidos_cantidad.txtPrecio = Me.lvDetalle.SelectedItem.SubItems(2)
frmFormGenPedidos_cantidad.Show vbModal
If frmFormGenPedidos_cantidad.gacepta Then
    Me.lvDetalle.SelectedItem.Text = frmFormGenPedidos_cantidad.gCantidad
    Me.lvDetalle.SelectedItem.SubItems(2) = frmFormGenPedidos_cantidad.gPrecio
    Me.lvDetalle.SelectedItem.SubItems(3) = frmFormGenPedidos_cantidad.gPrecio * frmFormGenPedidos_cantidad.gCantidad
    CalculoTotal
End If
End Sub

Private Sub lvDetalle_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdDel.Enabled = True
End Sub

Private Sub txtCantidad_Change()
ValidarSoloNumerosPunto Me.txtCantidad
CalculoImporte
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumerosPunto(Me.txtCantidad, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        Me.txtPrecio.SetFocus
        Me.txtPrecio.SelStart = 0
        Me.txtPrecio.SelLength = Len(Me.txtPrecio.Text)
    End If

End Sub


Private Sub txtPrecio_Change()
ValidarSoloNumerosPunto Me.txtPrecio
CalculoImporte
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumerosPunto(Me.txtPrecio, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        cmdAdd_Click

    End If
End Sub

Private Sub txtProducto_Change()
vBuscar = True
End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > lvSearch.ListItems.count Then loc_key = lvSearch.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > lvSearch.ListItems.count Then loc_key = lvSearch.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 27 Then
        Me.lvSearch.Visible = False
        Me.txtProducto.Text = ""
        Me.lblidprod.Caption = ""
        Me.txtCantidad.Text = ""
        Me.txtPrecio.Text = ""
        
        
    End If

    GoTo fin
POSICION:
    lvSearch.ListItems.Item(loc_key).Selected = True
    lvSearch.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(lvSearch.ListItems.Item(loc_key).Text) & " "
    txtRS.SelStart = Len(txtRS.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.lvSearch.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[dbo].[USP_PRODUCTO_SEARCH]"
            Set oRsPago = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtProducto.Text))

            Dim Item As Object
        
            If Not oRsPago.EOF Then

                Do While Not oRsPago.EOF
                    Set Item = Me.lvSearch.ListItems.Add(, , oRsPago!cod)
                    Item.SubItems(1) = Trim(oRsPago!prod)
                    Item.SubItems(2) = oRsPago!PRE
                    Item.Tag = oRsPago!ide
                    oRsPago.MoveNext
                Loop

                Me.lvSearch.Visible = True
                Me.lvSearch.ListItems(1).Selected = True
                loc_key = 1
                Me.lvSearch.ListItems(1).EnsureVisible
                vBuscar = False
          
            End If
        
        Else
            
            Me.lblidprod.Caption = Me.lvSearch.ListItems(loc_key).Tag
            Me.txtProducto.Tag = Me.lvSearch.ListItems(loc_key).Text
            Me.txtPrecio.Text = Me.lvSearch.ListItems(loc_key).SubItems(2)
            Me.txtProducto.Text = Me.lvSearch.ListItems(loc_key).SubItems(1)
            Me.txtCantidad.SetFocus
            Me.lvSearch.Visible = False
        End If
    End If
End Sub

Private Sub CalculoImporte()
If IsNumeric(Me.txtCantidad.Text) And IsNumeric(Me.txtPrecio.Text) Then
    Me.lblImporte.Caption = val(Me.txtCantidad.Text) * val(Me.txtPrecio.Text)
    Else
    Me.lblImporte.Caption = "0.00"
End If
End Sub

Private Sub CalculoTotal()
Dim itemx As Object
Dim vTotal As Double
vTotal = 0
For Each itemx In Me.lvDetalle.ListItems
    vTotal = vTotal + itemx.SubItems(3)
Next
Me.lblTotal.Caption = vTotal
End Sub
