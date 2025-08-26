VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmVDProductoBonificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Bonificaciones - Venta Directa"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13065
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
   ScaleHeight     =   7185
   ScaleWidth      =   13065
   Begin TabDlg.SSTab SSTTab0 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "Listado de Articulos"
      TabPicture(0)   =   "frmVDProductoBonificacion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvArticulos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Asignar Bonificación"
      TabPicture(1)   =   "frmVDProductoBonificacion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74760
         TabIndex        =   24
         Top             =   5760
         Width           =   12495
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   720
            Left            =   11040
            Picture         =   "frmVDProductoBonificacion.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "&Grabar"
            Height          =   720
            Left            =   8160
            Picture         =   "frmVDProductoBonificacion.frx":07A2
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "Eliminar"
            Enabled         =   0   'False
            Height          =   720
            Left            =   9600
            Picture         =   "frmVDProductoBonificacion.frx":0F0C
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   18
         Top             =   1800
         Width           =   12495
         Begin MSDataListLib.DataCombo DatCategoria 
            Height          =   360
            Left            =   6360
            TabIndex        =   7
            Top             =   1920
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.TextBox txtTope 
            Height          =   360
            Left            =   10320
            TabIndex        =   8
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton cmdBoniDel 
            Enabled         =   0   'False
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
            Left            =   11650
            Picture         =   "frmVDProductoBonificacion.frx":1296
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2880
            Width           =   750
         End
         Begin VB.CommandButton cmdBoniAdd 
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
            Left            =   11650
            Picture         =   "frmVDProductoBonificacion.frx":1620
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2400
            Width           =   750
         End
         Begin MSComctlLib.ListView lvBonificacion 
            Height          =   1455
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2400
            Width           =   11535
            _ExtentX        =   20346
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
         Begin VB.TextBox txtrecibe 
            Height          =   360
            Left            =   120
            TabIndex        =   5
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtCantidad 
            Height          =   360
            Left            =   1200
            TabIndex        =   4
            Top             =   1200
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo DatBonificacion 
            Height          =   360
            Left            =   1320
            TabIndex        =   6
            Top             =   1920
            Width           =   4935
            _ExtentX        =   8705
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
         Begin MSMask.MaskEdBox mebBIni 
            Height          =   375
            Left            =   3480
            TabIndex        =   2
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mebBFin 
            Height          =   375
            Left            =   6720
            TabIndex        =   3
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria:"
            Height          =   240
            Left            =   6360
            TabIndex        =   29
            Top             =   1680
            Width           =   1035
         End
         Begin VB.Line Line2 
            BorderStyle     =   2  'Dash
            X1              =   240
            X2              =   11280
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vigencia:"
            Height          =   240
            Left            =   1320
            TabIndex        =   28
            Top             =   660
            Width           =   915
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   240
            Left            =   2640
            TabIndex        =   27
            Top             =   660
            Width           =   690
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   240
            Left            =   6000
            TabIndex        =   26
            Top             =   660
            Width           =   660
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tope Boni:"
            Height          =   240
            Left            =   10320
            TabIndex        =   25
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label lblProducto2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2400
            TabIndex        =   23
            Tag             =   "X"
            Top             =   1200
            Width           =   9015
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recibe:"
            Height          =   240
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Por cada:"
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   1260
            Width           =   960
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Producto:"
            Height          =   240
            Left            =   1320
            TabIndex        =   20
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ASIGNAR BONIFICACIÓN"
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2385
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   15
         Top             =   480
         Width           =   12495
         Begin VB.Label lblIdProducto 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11280
            TabIndex        =   17
            Tag             =   "X"
            Top             =   120
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lblProducto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   675
            Left            =   360
            TabIndex        =   16
            Tag             =   "X"
            Top             =   360
            Width           =   11715
         End
      End
      Begin MSComctlLib.ListView lvArticulos 
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   11245
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
   End
End
Attribute VB_Name = "frmVDProductoBonificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pIDempresa As Integer
Private vPUNTO As Boolean 'variable para controld epunto sin utilizar ocx
Private oRSboni As New ADODB.Recordset

Private Function ObtenerMaximoIDBoni(rs As ADODB.Recordset) As Integer

    On Error GoTo ErrorHandler
    
    Dim maxID As Integer

    maxID = 0 ' Valor inicial
    
    ' Verificar que el Recordset esté abierto y tenga registros
    If rs.State = adStateClosed Then
        MsgBox "El Recordset está cerrado", vbExclamation
        Exit Function

    End If

    ' Guardar posición actual
    Dim bookmark As Variant

    If Not rs.EOF Then
        bookmark = rs.bookmark
    
        ' Buscar el máximo valor de idprecio
        rs.MoveFirst

        Do Until rs.EOF

            If Not IsNull(rs!idboni.Value) Then
                If rs!idboni.Value > maxID Then
                    maxID = rs!idboni.Value

                End If

            End If

            rs.MoveNext
        Loop
    
        ' Restaurar posición original
        rs.bookmark = bookmark

    End If

    Dim itemx      As Object

    Dim IdBoniMaxLV As Integer

    IdBoniMaxLV = 0

    For Each itemx In Me.lvBonificacion.ListItems

        If itemx.Tag > IdBoniMaxLV Then
            IdBoniMaxLV = itemx.Tag

        End If

    Next

    If IdBoniMaxLV > maxID Then
        ObtenerMaximoIDBoni = IdBoniMaxLV + 1
    Else
        ObtenerMaximoIDBoni = maxID + 1

    End If

    Exit Function
    
ErrorHandler:
    MsgBox "Error al buscar máximo ID: " & Err.Description, vbExclamation
    ObtenerMaximoIDBoni = -1 ' Valor de error

End Function

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
    Dim Item      As Object

    Dim idMaxBoni As Integer

    Dim cruce     As Boolean

    cruce = False
    
    If Me.lvBonificacion.ListItems.count = 0 Then
        If oRSboni.RecordCount = 0 Then
            idMaxBoni = 1
        Else
            idMaxBoni = ObtenerMaximoIDBoni(oRSboni)

        End If

        Set Item = Me.lvBonificacion.ListItems.Add(, , Me.txtCantidad.Text)
        Item.Tag = idMaxBoni
        Item.SubItems(1) = Me.lblProducto2.Caption ' Me.DatBonificacion.BoundText
        Item.SubItems(2) = Me.DatCategoria.BoundText ' Me.DatBonificacion.Text
        Item.SubItems(3) = Me.DatCategoria.Text ' Me.txtrecibe.Text
        Item.SubItems(4) = Me.txtrecibe.Text
        Item.SubItems(5) = Me.DatBonificacion.BoundText ' .txtTope.Text
        Item.SubItems(6) = Me.DatBonificacion.Text ' Me.DatCategoria.BoundText
        Item.SubItems(7) = Me.txtTope.Text
    Else

        For Each itemx In Me.lvBonificacion.ListItems

            If Me.DatCategoria.BoundText = itemx.SubItems(2) Then
                cruce = True
                Exit For

            End If

        Next
        
        If cruce Then
            MsgBox "Producto ya se encuentra Agregado.", vbCritical, Pub_Titulo
            Me.DatBonificacion.SetFocus
            Exit Sub
        Else
            idMaxBoni = ObtenerMaximoIDBoni(oRSboni)
            Set Item = Me.lvBonificacion.ListItems.Add(, , Me.txtCantidad.Text)
            Item.Tag = idMaxBoni
            Item.SubItems(1) = Me.lblProducto2.Caption ' Me.DatBonificacion.BoundText
            Item.SubItems(2) = Me.DatCategoria.BoundText ' Me.DatBonificacion.Text
            Item.SubItems(3) = Me.DatCategoria.Text ' Me.txtrecibe.Text
            Item.SubItems(4) = Me.txtrecibe.Text
            Item.SubItems(5) = Me.DatBonificacion.BoundText ' .txtTope.Text
            Item.SubItems(6) = Me.DatBonificacion.Text ' Me.DatCategoria.BoundText
            Item.SubItems(7) = Me.txtTope.Text

        End If

    End If
    
    Me.txtCantidad.Text = ""
    Me.DatBonificacion.BoundText = -1
    Me.DatCategoria.BoundText = -1
    Me.txtrecibe.Text = ""
    Me.txtTope.Text = ""
    
    Me.txtCantidad.SetFocus

End Sub

Private Sub cmdBoniDel_Click()

    If Me.lvBonificacion.ListItems.count = 0 Then Exit Sub
    If Me.lvBonificacion.SelectedItem Is Nothing Then Exit Sub

    With Me.lvBonificacion.SelectedItem
        oRSboni.AddNew
        oRSboni.Fields(0).Value = .Tag
        oRSboni.Update

    End With

    Me.lvBonificacion.ListItems.Remove Me.lvBonificacion.SelectedItem.Index
    Me.cmdBoniDel.Enabled = False
    Me.txtCantidad.SetFocus

End Sub

Private Sub cmdEliminar_Click()

    If MsgBox("¿Desea eliminar toda la configuración" + vbCrLf + "del producto " + Me.lblProducto.Caption + "?", vbQuestion + vbYesNo, "Eliminar Configuración") = vbNo Then Exit Sub

    On Error GoTo cElimina

    MousePointer = vbHourglass
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_BONIFICACION_DELETE]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lblIdProducto.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adBSTR, adParamInput, 20, LK_CODUSU)
    
    Dim orsResult As ADODB.Recordset

    Set orsResult = oCmdEjec.Execute

    Dim Smensaje() As String

    If Not orsResult.EOF Then
        Smensaje = Split(orsResult.Fields(0), "=")

        If Smensaje(0) = 0 Then
              
            MousePointer = vbDefault
            MsgBox Smensaje(1), vbInformation, Pub_Titulo
            cmdCancelar_Click
        
        Else
            MousePointer = vbDefault
            MsgBox Smensaje(1), vbCritical, Pub_Titulo

        End If

    End If

    CerrarConexion True
    Exit Sub
cElimina:
    CerrarConexion True
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub cmdGrabar_Click()

    If Len(Trim(Me.lblProducto.Caption)) = 0 Then
        MsgBox "Debe elegir un producto para continuar.", vbInformation, Pub_Titulo
        Exit Sub

    End If
    
    If Me.lvBonificacion.ListItems.count <> 0 And Not ValidarFecha(Me.mebBIni.Text) Then
        MsgBox "Debe ingresar la Fecha Inicial de la Bonificacion.", vbInformation, Pub_Titulo
        Me.mebBIni.SetFocus
        Exit Sub

    End If

    If Me.lvBonificacion.ListItems.count <> 0 And Not ValidarFecha(Me.mebBIni.Text) Then
        MsgBox "Debe ingresar la Fecha Inicial de la Bonificacion.", vbInformation, Pub_Titulo
        Me.mebBIni.SetFocus
        Exit Sub

    End If

    On Error GoTo cGraba

    MousePointer = vbHourglass
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_BONIFICACION_PROCCESS]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lblIdProducto.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)
    
    'OBTENIENDO XML DE PROMOCIONES
    Dim itemx           As Object
    
    'OBTENIENDO XML DE BONIFICACIONES
    Dim strBonificacion As String

    If Me.lvBonificacion.ListItems.count <> 0 Then
        strBonificacion = "<r>"

        For Each itemx In Me.lvBonificacion.ListItems

            strBonificacion = strBonificacion & "<d "
            strBonificacion = strBonificacion & "idboni=""" & itemx.Tag & """ "
            strBonificacion = strBonificacion & "cantbase=""" & itemx.Text & """ "
            strBonificacion = strBonificacion & "idpboni=""" & itemx.SubItems(5) & """ "
            strBonificacion = strBonificacion & "cantboni=""" & itemx.SubItems(4) & """ "
            strBonificacion = strBonificacion & "tope=""" & IIf(Len(Trim(itemx.SubItems(7))) = 0, "", itemx.SubItems(7)) & """ "
            strBonificacion = strBonificacion & "idcat=""" & itemx.SubItems(2) & """ "
            strBonificacion = strBonificacion & "/>"
        Next
        strBonificacion = strBonificacion & "</r>"
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xBONIFICACION", adVarChar, adParamInput, 4000, strBonificacion)
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FINI", adVarChar, adParamInput, 8, ConvertirFechaFormat_yyyyMMdd(Me.mebBIni.Text))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FFIN", adVarChar, adParamInput, 8, ConvertirFechaFormat_yyyyMMdd(Me.mebBFin.Text))

    End If
    
    '    oCmdEjec.Execute
    Dim orsResult As ADODB.Recordset

    Set orsResult = oCmdEjec.Execute

    Dim Smensaje() As String

    If Not orsResult.EOF Then
        Smensaje = Split(orsResult.Fields(0), "=")

        If Smensaje(0) = 0 Then
          
            MousePointer = vbDefault
            MsgBox Smensaje(1), vbInformation, Pub_Titulo
            cmdCancelar_Click
        Else
            MousePointer = vbDefault
            MsgBox Smensaje(1), vbCritical, Pub_Titulo

        End If

    End If

    CerrarConexion True
    Exit Sub
cGraba:
    CerrarConexion True
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub



Private Sub cmdCancelar_Click()
LimpiarControles Me
EliminarRegistrosRecordSet oRSboni
Me.lvBonificacion.ListItems.Clear
Me.SSTTab0.tab = 0
End Sub


Private Sub DatBonificacion_KeyDown(KeyCode As Integer, Shift As Integer)
HandleEnterKey KeyCode, Me.DatCategoria
End Sub

Private Sub DatCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
HandleEnterKey KeyCode, Me.txtTope
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
If oRSboni.State = adStateOpen Then oRSboni.Close
    oRSboni.CursorLocation = adUseClient
    oRSboni.Fields.Append "idboni", adInteger
    oRSboni.Open
    
pIDempresa = devuelveIDempresaXdefecto
CentrarFormulario MDIForm1, Me
ConfigurarLV
cargarProductos
End Sub

Private Sub cargarProductos()
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_LIST]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@idempresa", adInteger, adParamInput, , pIDempresa)
  
    Dim orsData As ADODB.Recordset

    Set orsData = oCmdEjec.Execute
  
    Dim itemx As Object
  
    Do While Not orsData.EOF
        Set itemx = Me.lvArticulos.ListItems.Add(, , orsData!ide)
        itemx.SubItems(1) = orsData!nom
        orsData.MoveNext
  
    Loop
    CerrarConexion True
  
End Sub

Private Sub cargarProductosCombo()
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_LIST]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@LISTADO", adBoolean, adParamInput, , True)

    Set oRSmain = oCmdEjec.Execute
    
    Dim orsTEMP1 As New ADODB.Recordset
    
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
  
    Set Me.DatBonificacion.RowSource = orsTEMP1
    Me.DatBonificacion.ListField = orsTEMP1(1).Name
    Me.DatBonificacion.BoundColumn = orsTEMP1(0).Name
    Me.DatBonificacion.BoundText = -1
    
    Set oRSmain = oRSmain.NextRecordset
    Dim orsTEMP2 As New ADODB.Recordset
    
    orsTEMP2.CursorLocation = adUseClient
    orsTEMP2.Fields.Append oRSmain(0).Name, oRSmain(0).Type, oRSmain(0).DefinedSize
    orsTEMP2.Fields.Append oRSmain(1).Name, oRSmain(1).Type, oRSmain(1).DefinedSize
    orsTEMP2.Open
    
    oRSmain.MoveFirst
    
    Do Until oRSmain.EOF
        orsTEMP2.AddNew
        orsTEMP2.Fields(0).Value = oRSmain.Fields(0).Value
        orsTEMP2.Fields(1).Value = oRSmain.Fields(1).Value
        orsTEMP2.Update
        oRSmain.MoveNext
    Loop
    
    Set Me.DatCategoria.RowSource = orsTEMP2
    Me.DatCategoria.ListField = orsTEMP2(1).Name
    Me.DatCategoria.BoundColumn = orsTEMP2(0).Name
    Me.DatCategoria.BoundText = -1
    
    CerrarConexion True

End Sub

Private Sub ConfigurarLV()

    With Me.lvArticulos
        .ColumnHeaders.Add , , "Código", 1500
        .ColumnHeaders.Add , , "Producto", 7000
        .FullRowSelect = True
        .Gridlines = True
        .HideColumnHeaders = False
        .View = lvwReport
        .HideSelection = False

    End With

    With Me.lvBonificacion
        .ColumnHeaders.Add , , "Por cada"
        .ColumnHeaders.Add , , "Producto", 2500
        .ColumnHeaders.Add , , "idCategoria", 0
        .ColumnHeaders.Add , , "Categoria", 1800
         .ColumnHeaders.Add , , "Recibe"
        .ColumnHeaders.Add , , "idpBoni", 0
        .ColumnHeaders.Add , , "Bonificación", 2500
        .ColumnHeaders.Add , , "Tope", 900

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
    EliminarRegistrosRecordSet oRSboni
    Me.lblProducto.Caption = Me.lvArticulos.SelectedItem.SubItems(1)
    Me.lblIdProducto.Caption = Me.lvArticulos.SelectedItem.Text
    Me.lblProducto2.Caption = Me.lvArticulos.SelectedItem.SubItems(1)
    cargarProductosCombo
    obtenerInformacionPromocion Me.lvArticulos.SelectedItem.Text
    Me.mebBIni.SetFocus
End Sub

Private Sub lvBonificacion_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdBoniDel.Enabled = True
End Sub


Private Sub mebBFin_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.txtCantidad
End Sub

Private Sub mebBIni_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.mebBFin
End Sub


Private Sub txtCantidad_Change()
ValidarSoloNumeros Me.txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
 KeyAscii = SoloNumeros(KeyAscii)
 HandleEnterKey KeyAscii, Me.txtrecibe
End Sub



Private Sub txtrecibe_Change()
ValidarSoloNumeros Me.txtrecibe
End Sub

Private Sub txtrecibe_KeyPress(KeyAscii As Integer)

   KeyAscii = SoloNumeros(KeyAscii)
    HandleEnterKey KeyAscii, Me.DatBonificacion

End Sub

Private Sub obtenerInformacionPromocion(cIDProducto As Integer)
    Me.lvBonificacion.ListItems.Clear
    LimpiarMaskEdBox "##/##/####", Me.mebBIni
    LimpiarMaskEdBox "##/##/####", Me.mebBFin

    On Error GoTo cDatos

    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_BONIFICACION_FILL]"
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , cIDProducto)
  
   

    Set oRSmain = oCmdEjec.Execute
  
    Dim itemx As Object
    
    
    Do While Not oRSmain.EOF
        Set itemx = Me.lvBonificacion.ListItems.Add(, , oRSmain!cant_base)
        itemx.Tag = oRSmain!idboni
        itemx.SubItems(1) = Me.lblProducto2.Caption ' oRSmain!idpboni
        itemx.SubItems(2) = oRSmain!idcat ' oRSmain!producto
        itemx.SubItems(3) = oRSmain!nomcat ' oRSmain!cant_boni
                itemx.SubItems(4) = oRSmain!cant_boni
        itemx.SubItems(5) = oRSmain!idpboni
        itemx.SubItems(6) = oRSmain!producto

        itemx.SubItems(7) = oRSmain!tope
        If Not IsNull(oRSmain!FINI) Then Me.mebBIni.Text = Nulo_Valors(oRSmain!FINI)
        If Not IsNull(oRSmain!ffin) Then Me.mebBFin.Text = Nulo_Valors(oRSmain!ffin)
        oRSmain.MoveNext
    Loop
    
    Set oRSmain = oRSmain.NextRecordset
    
    If Not oRSmain.EOF Then
    Do While Not oRSmain.EOF
          oRSboni.AddNew
        oRSboni.Fields(0).Value = oRSmain!idboni
        oRSboni.Update
        oRSmain.MoveNext
        Loop
    End If
    
  
    If Me.lvBonificacion.ListItems.count <> 0 Then
        Me.cmdEliminar.Enabled = True
    Else
        Me.cmdEliminar.Enabled = False

    End If
  CerrarConexion True
    Exit Sub
cDatos:
CerrarConexion True
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub txtTope_Change()
ValidarSoloNumeros Me.txtTope
End Sub

Private Sub txtTope_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
 If KeyAscii = vbKeyReturn Then cmdBoniAdd_Click
End Sub
