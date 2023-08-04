VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmProductoPromocion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Promociones"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13065
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   13065
   Begin TabDlg.SSTab SSTTab0 
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado de Articulos"
      TabPicture(0)   =   "frmProductoPromocion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvArticulos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Asignar Promoción"
      TabPicture(1)   =   "frmProductoPromocion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   -74760
         TabIndex        =   28
         Top             =   8400
         Width           =   12495
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   480
            Left            =   11280
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "&Grabar"
            Height          =   480
            Left            =   10080
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   16
         Top             =   5160
         Width           =   12495
         Begin VB.CommandButton cmdBoniDel 
            Enabled         =   0   'False
            Height          =   360
            Left            =   11160
            TabIndex        =   27
            Top             =   2160
            Width           =   990
         End
         Begin VB.CommandButton cmdBoniAdd 
            Height          =   360
            Left            =   11160
            TabIndex        =   26
            Top             =   1680
            Width           =   990
         End
         Begin MSComctlLib.ListView lvBonificacion 
            Height          =   1455
            Left            =   120
            TabIndex        =   24
            Top             =   1680
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   2566
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
         Begin VB.TextBox txtrecibe 
            Height          =   375
            Left            =   9120
            TabIndex        =   23
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtCantidad 
            Height          =   375
            Left            =   9120
            TabIndex        =   21
            Top             =   480
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo DatBonificacion 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblProducto2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Tag             =   "X"
            Top             =   480
            Width           =   8775
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recibe:"
            Height          =   195
            Left            =   9120
            TabIndex        =   22
            Top             =   960
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Por cada:"
            Height          =   195
            Left            =   9240
            TabIndex        =   20
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Producto:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ASIGNAR BONIFICACIÓN"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   2205
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   5
         Top             =   1800
         Width           =   12495
         Begin VB.CommandButton cmdPromDel 
            Enabled         =   0   'False
            Height          =   360
            Left            =   11400
            TabIndex        =   15
            Top             =   1680
            Width           =   990
         End
         Begin MSComctlLib.ListView lvPromocion 
            Height          =   2175
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   3836
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
            Left            =   10080
            TabIndex        =   12
            Top             =   555
            Width           =   975
         End
         Begin VB.TextBox txtHasta 
            Height          =   285
            Left            =   5400
            TabIndex        =   11
            Top             =   555
            Width           =   975
         End
         Begin VB.TextBox txtDesde 
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdPromAdd 
            Height          =   360
            Left            =   11400
            TabIndex        =   14
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Precio:"
            Height          =   195
            Left            =   9360
            TabIndex        =   9
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rango Final:"
            Height          =   195
            Left            =   4200
            TabIndex        =   8
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rango Inicial:"
            Height          =   195
            Left            =   360
            TabIndex        =   7
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ASIGNAR PROMOCIÓN"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1965
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   12495
         Begin VB.Label lblIdProducto 
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   11280
            TabIndex        =   4
            Tag             =   "X"
            Top             =   120
            Width           =   555
         End
         Begin VB.Label lblProducto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   675
            Left            =   360
            TabIndex        =   3
            Tag             =   "X"
            Top             =   360
            Width           =   11715
         End
      End
      Begin MSComctlLib.ListView lvArticulos 
         Height          =   8775
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   15478
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
End
Attribute VB_Name = "frmProductoPromocion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPUNTO As Boolean 'variable para controld epunto sin utilizar ocx

Private Sub cmdBoniAdd_Click()

    If Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Debe ingresar el valor [Por cada].", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Exit Sub

    End If

    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "Valor incorrecto [Por cada].", vbInformation, Pub_Titulo
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
        Me.txtCantidad.SetFocus
        Exit Sub

    End If

    If Me.DatBonificacion.BoundText = -1 Then
        MsgBox "Debe elegir el Producto Bonificación.", vbInformation, Pub_Titulo
        Me.DatBonificacion.SetFocus
        Exit Sub

    End If

    If Len(Trim(Me.txtrecibe.Text)) = 0 Then
        MsgBox "Debe ingresar el valor [Recibe].", vbCritical, Pub_Titulo
        Me.txtrecibe.SetFocus
        Exit Sub

    End If

    If Not IsNumeric(Me.txtrecibe.Text) Then
        MsgBox "Valor incorrecto [Recibe].", vbInformation, Pub_Titulo
        Me.txtrecibe.SelStart = 0
        Me.txtrecibe.SelLength = Len(Me.txtrecibe.Text)
        Me.txtrecibe.SetFocus
        Exit Sub

    End If

    'agregar item
    Dim Item As Object
    
    If Me.lvBonificacion.ListItems.count = 0 Then
        Set Item = Me.lvBonificacion.ListItems.Add(, , Me.txtCantidad.Text)
        Item.SubItems(1) = Me.DatBonificacion.BoundText
        Item.SubItems(2) = Me.DatBonificacion.Text
        Item.SubItems(3) = Me.txtrecibe.Text
        Item.SubItems(4) = "0.00"
    Else

        Dim itemx As Object

        Dim cruce As Boolean

        cruce = False

        For Each itemx In Me.lvBonificacion.ListItems

            If Me.DatBonificacion.BoundText = itemx.SubItems(1) Then
                cruce = True
                Exit For

            End If

        Next

        If cruce = False Then
            Set Item = Me.lvBonificacion.ListItems.Add(, , Me.txtCantidad.Text)
            Item.SubItems(1) = Me.DatBonificacion.BoundText
            Item.SubItems(2) = Me.DatBonificacion.Text
            Item.SubItems(3) = Me.txtrecibe.Text
            Item.SubItems(4) = "0.00"
        Else
            MsgBox "Producto ya se encuentra Agregado.", vbCritical, Pub_Titulo
            Me.DatBonificacion.SetFocus
            Exit Sub

        End If

    End If
    
    Me.txtCantidad.Text = ""
    Me.DatBonificacion.BoundText = -1
    Me.txtrecibe.Text = ""
    
    Me.txtCantidad.SetFocus

End Sub

Private Sub cmdBoniDel_Click()
If Me.lvBonificacion.SelectedItem Is Nothing Then Exit Sub
Me.lvBonificacion.ListItems.Remove Me.lvBonificacion.SelectedItem.Index
Me.cmdBoniDel.Enabled = False
Me.txtCantidad.SetFocus
End Sub

Private Sub cmdGrabar_Click()

    If Me.lvPromocion.ListItems.count = 0 Then
        MsgBox "Debe ingresar promociones", vbCritical, Pub_Titulo
        Exit Sub

    End If

    On Error GoTo cGraba

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_PROMOCION_BONIFICACION_PROCCESS]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lblIdProducto.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRODUCTO", adBSTR, adParamInput, 200, Me.lblProducto.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adBSTR, adParamInput, 20, LK_CODUSU)
    
    'OBTENIENDO XML DE PROMOCIONES
    Dim itemx        As Object

    Dim strPromocion As String

    If Me.lvPromocion.ListItems.count <> 0 Then
        strPromocion = "<r>"

        For Each itemx In Me.lvPromocion.ListItems

            strPromocion = strPromocion & "<d "
            strPromocion = strPromocion & "idprecio=""" & itemx.Tag & """ "
            strPromocion = strPromocion & "ini=""" & itemx.SubItems(1) & """ "
            strPromocion = strPromocion & "fin=""" & itemx.SubItems(2) & """ "
            strPromocion = strPromocion & "pre=""" & itemx.SubItems(3) & """ "
            strPromocion = strPromocion & "/>"
            
        Next
        strPromocion = strPromocion & "</r>"

    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xPROMOCION", adBSTR, adParamInput, 4000, strPromocion)
    
    'OBTENIENDO XML DE BONIFICACIONES
    Dim strBonificacion As String

    If Me.lvBonificacion.ListItems.count <> 0 Then
        strBonificacion = "<r>"

        For Each itemx In Me.lvBonificacion.ListItems

            strBonificacion = strBonificacion & "<d "
            strBonificacion = strBonificacion & "idboni=""" & itemx.SubItems(1) & """ "
            strBonificacion = strBonificacion & "cant=""" & itemx.Text & """ "
            strBonificacion = strBonificacion & "boni=""" & itemx.SubItems(3) & """ "
            strBonificacion = strBonificacion & "pre=""" & itemx.SubItems(4) & """ "
            strBonificacion = strBonificacion & "/>"
        Next
        strBonificacion = strBonificacion & "</r>"
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xBONIFICACION", adBSTR, adParamInput, 4000, strBonificacion)

    End If
    
'    oCmdEjec.Execute
    Dim orsResult  As ADODB.Recordset
    Set orsResult = oCmdEjec.Execute
    Dim sMensaje() As String
    If Not orsResult.EOF Then
        sMensaje = Split(orsResult.Fields(0), "=")
        If sMensaje(0) = 0 Then
        MsgBox sMensaje(1), vbInformation, Pub_Titulo
    cmdCancelar_Click
        Else
        MsgBox sMensaje(1), vbCritical, Pub_Titulo
        End If
    End If
    
    
    Exit Sub
cGraba:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub cmdPromAdd_Click()

    If Len(Trim(Me.txtDesde.Text)) = 0 Then
        MsgBox "Debe ingresar el Rango Inicial.", vbCritical, Pub_Titulo
        Me.txtDesde.SetFocus
        Exit Sub

    End If

    '    If Len(Trim(Me.txtHasta.Text)) = 0 Then
    '        MsgBox "Debe ingresar el Rango Final.", vbCritical, Pub_Titulo
    '        Me.txtHasta.SetFocus
    '        Exit Sub
    '    End If

    If Not IsNumeric(Me.txtDesde.Text) Then
        MsgBox "Rango Inicial incorrecto.", vbCritical, Pub_Titulo
        Me.txtDesde.SetFocus
        Me.txtDesde.SelStart = 0
        Me.txtDesde.SelLength = Len(Me.txtDesde.Text)

    End If

    '    If Not IsNumeric(Me.txtHasta.Text) Then
    '        MsgBox "Rango Final incorrecto.", vbCritical, Pub_Titulo
    '        Me.txtHasta.SetFocus
    '        Me.txtHasta.SelStart = 0
    '        Me.txtHasta.SelLength = Len(Me.txtHasta.Text)
    '        Exit Sub
    '    End If
    '
    If Len(Trim(Me.txtPrecio.Text)) = 0 Then
        MsgBox "Debe ingresar el precio.", vbCritical, Pub_Titulo
        Me.txtPrecio.SetFocus
        Exit Sub

    End If
    
    If Val(Me.txtPrecio.Text) <= 0 Then
        MsgBox "El precio ingresado es incorrecto.", vbInformation, Pub_Titulo
        Me.txtPrecio.SetFocus
        Me.txtPrecio.SelStart = 0
        Me.txtPrecio.SelLength = Len(Me.txtPrecio.Text)
        Exit Sub

    End If

    If IsNumeric(Me.txtHasta.Text) Then
        If Val(Me.txtDesde.Text) > Val(Me.txtHasta.Text) Then
            MsgBox "La cantidad inicial debe ser anterior al rango final", vbInformation, Pub_Titulo
            Exit Sub

        End If

    End If

    'agregar item
    Dim Item As Object

    If Me.lvPromocion.ListItems.count = 0 Then
        Set Item = Me.lvPromocion.ListItems.Add(, , Me.lblProducto.Caption)
        Item.Tag = Me.lvPromocion.ListItems.count
        Item.SubItems(1) = Me.txtDesde.Text
        Item.SubItems(2) = Me.txtHasta.Text
        Item.SubItems(3) = Me.txtPrecio.Text
    Else

        Dim itemx As Object

        Dim cruce As Boolean

        cruce = False

        For Each itemx In Me.lvPromocion.ListItems

            If Me.txtDesde.Text >= Val(itemx.SubItems(1)) And Me.txtDesde.Text <= Val(itemx.SubItems(2)) Then
                cruce = True
                Exit For

            End If

        Next

        If cruce = False Then
            Set Item = Me.lvPromocion.ListItems.Add(, , Me.lblProducto.Caption)
            Item.Tag = Me.lvPromocion.ListItems.count
            Item.SubItems(1) = Me.txtDesde.Text
            Item.SubItems(2) = Me.txtHasta.Text
            Item.SubItems(3) = Me.txtPrecio.Text
        Else
            MsgBox "Rangos se cruzan con otros ya ingresados.", vbCritical, Pub_Titulo
            Me.txtDesde.SetFocus
            Me.txtDesde.SelStart = 0
            Me.txtDesde.SelLength = Len(Me.txtDesde.Text)
            Exit Sub

        End If

    End If
    
    Me.txtDesde.Text = ""
    Me.txtHasta.Text = ""
    Me.txtPrecio.Text = ""
    
    Me.txtDesde.SetFocus

End Sub

Private Sub cmdCancelar_Click()
LimpiarControles Me
Me.lvBonificacion.ListItems.Clear
Me.lvPromocion.ListItems.Clear
Me.SSTTab0.tab = 0
End Sub

Private Sub cmdPromDel_Click()
If Me.lvPromocion.SelectedItem Is Nothing Then Exit Sub
Me.lvPromocion.ListItems.Remove Me.lvPromocion.SelectedItem.Index
Me.cmdPromDel.Enabled = False
Me.txtDesde.SetFocus
End Sub

Private Sub DatBonificacion_Change()
Me.txtrecibe.SetFocus
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
ConfigurarLV
cargarProductos
End Sub

Private Sub cargarProductos()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_PRODUCTO_LIST]"
  oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
  
  Dim orsData As ADODB.Recordset
  Set orsData = oCmdEjec.Execute
  
  Dim itemx As Object
  
  Do While Not orsData.EOF
  Set itemx = Me.lvArticulos.ListItems.Add(, , orsData!cod)
  itemx.SubItems(1) = orsData!nom
    orsData.MoveNext
  
  Loop
  
End Sub

Private Sub cargarProductosCombo()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_PRODUCTO_LIST]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@LISTADO", adBoolean, adParamInput, , True)
  
    Dim orsData As ADODB.Recordset

    Set orsData = oCmdEjec.Execute
  
    Set Me.DatBonificacion.RowSource = orsData
    Me.DatBonificacion.ListField = orsData(1).Name
    Me.DatBonificacion.BoundColumn = orsData(0).Name
    Me.DatBonificacion.BoundText = -1

End Sub

Private Sub ConfigurarLV()
With Me.lvArticulos
    .ColumnHeaders.Add , , "Código", 1500
    .ColumnHeaders.Add , , "Producto", 5000
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = False
    .View = lvwReport
    .HideSelection = False
End With

With Me.lvPromocion
    .ColumnHeaders.Add , , "Producto", 4500
    .ColumnHeaders.Add , , "Rango Inicial", 2000
    .ColumnHeaders.Add , , "Rango Final", 2000
    .ColumnHeaders.Add , , "Precio", 1000
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = False
    .View = lvwReport
    .HideSelection = False
End With

With Me.lvBonificacion
    .ColumnHeaders.Add , , "Por cada"
    .ColumnHeaders.Add , , "idBono", 0
    .ColumnHeaders.Add , , "Bonificación", 5500
    .ColumnHeaders.Add , , "Recibe"
    .ColumnHeaders.Add , , "Precio"
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = False
    .View = lvwReport
    .HideSelection = False
    

End With
End Sub

Private Sub lvArticulos_DblClick()
EnviarDatos
End Sub

Private Sub EnviarDatos()
    Me.SSTTab0.tab = 1
    Me.lblProducto.Caption = Me.lvArticulos.SelectedItem.SubItems(1)
    Me.lblIdProducto.Caption = Me.lvArticulos.SelectedItem.Text
    Me.lblProducto2.Caption = Me.lvArticulos.SelectedItem.SubItems(1)
    cargarProductosCombo
    Me.lvPromocion.ListItems.Clear
    Me.txtDesde.SetFocus
    obtenerInformacionPromocion Me.lvArticulos.SelectedItem.Text
End Sub

Private Sub lvBonificacion_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdBoniDel.Enabled = True
End Sub

Private Sub lvPromocion_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdPromDel.Enabled = True
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
 If SoloNumeros(KeyAscii) Then KeyAscii = 0
 If KeyAscii = vbKeyReturn Then Me.DatBonificacion.SetFocus
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
If KeyAscii = vbKeyReturn Then Me.txtHasta.SetFocus
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
If KeyAscii = vbKeyReturn Then Me.txtPrecio.SetFocus
End Sub

Private Sub txtPrecio_Change()
If InStr(Me.txtPrecio.Text, ".") Then
    vPUNTO = True
Else
    vPUNTO = False
End If

End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)

    If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If KeyAscii = 46 Then
        If vPUNTO Or Len(Trim(Me.txtPrecio.Text)) = 0 Then
            KeyAscii = 0

        End If

    End If
    
    If KeyAscii = vbKeyReturn Then cmdPromAdd_Click

    
End Sub

Private Sub txtrecibe_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub obtenerInformacionPromocion(cIDProducto As Integer)

    On Error GoTo cDatos

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_PROMOCION_BONIFICACION_FILL]"
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , cIDProducto)
  
    Dim orsData As ADODB.Recordset

    Set orsData = oCmdEjec.Execute
  
    Dim itemx As Object
  
    Do While Not orsData.EOF
        Set itemx = Me.lvPromocion.ListItems.Add(, , orsData!descripcion)
        itemx.Tag = orsData!idepre
        itemx.SubItems(1) = orsData!ini
        itemx.SubItems(2) = IIf(orsData!fin = 0, "", orsData!fin)
        itemx.SubItems(3) = orsData!PRE
        orsData.MoveNext
    
    Loop
    
    Dim orsBoni As ADODB.Recordset
    Set orsBoni = orsData.NextRecordset
    
    Do While Not orsBoni.EOF
        Set itemx = Me.lvBonificacion.ListItems.Add(, , orsBoni!cant)
        itemx.SubItems(1) = orsBoni!idboni
        itemx.SubItems(2) = orsBoni!producto
        itemx.SubItems(3) = orsBoni!boni
        itemx.SubItems(4) = orsBoni!PRE
        orsBoni.MoveNext
    Loop
    
       
  
    Exit Sub
cDatos:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

