VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{FEC367D0-B73E-4DD0-80FD-1F56BC27B04A}#1.0#0"; "McToolBar.ocx"
Begin VB.Form frmMantVDProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Productos [Venta Directa]"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMantVDProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   11895
   Begin MSComctlLib.ImageList ilProducto 
      Left            =   11880
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVDProducto.frx":0CCA
            Key             =   "product"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTProducto 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmMantVDProducto.frx":1444
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtSearch"
      Tab(0).Control(1)=   "lvProducto"
      Tab(0).Control(2)=   "Label1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Producto"
      TabPicture(1)   =   "frmMantVDProducto.frx":1460
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FraProducto"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FraProducto 
         Height          =   7215
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   11055
         Begin VB.Frame FraPrecio 
            Height          =   4095
            Left            =   120
            TabIndex        =   17
            Top             =   3000
            Width           =   10815
            Begin VB.CommandButton cmdAdd 
               Height          =   360
               Left            =   10200
               Picture         =   "frmMantVDProducto.frx":147C
               Style           =   1  'Graphical
               TabIndex        =   23
               Tag             =   "X"
               Top             =   1200
               Width           =   495
            End
            Begin MSComctlLib.ListView lvPrecios 
               Height          =   2895
               Left            =   120
               TabIndex        =   22
               Tag             =   "X"
               Top             =   1080
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   5106
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
            Begin VB.TextBox txtAddPrecio 
               Height          =   360
               Left            =   8160
               MaxLength       =   6
               TabIndex        =   21
               Tag             =   "X"
               Top             =   600
               Width           =   1335
            End
            Begin MSDataListLib.DataCombo datAddCategoria 
               Height          =   360
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   7815
               _ExtentX        =   13785
               _ExtentY        =   635
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin VB.CommandButton cmdDel 
               Height          =   360
               Left            =   10200
               Picture         =   "frmMantVDProducto.frx":1806
               Style           =   1  'Graphical
               TabIndex        =   24
               Tag             =   "X"
               Top             =   1560
               Width           =   495
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Precio"
               Height          =   240
               Left            =   8160
               TabIndex        =   19
               Top             =   360
               Width           =   600
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Categoria"
               Height          =   240
               Left            =   120
               TabIndex        =   18
               Top             =   360
               Width           =   945
            End
         End
         Begin VB.Frame FraMain 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   10815
            Begin VB.TextBox txtCodAlternativo 
               Height          =   360
               Left            =   2640
               TabIndex        =   9
               Tag             =   "X"
               Top             =   240
               Width           =   3015
            End
            Begin VB.TextBox txtStock 
               Height          =   360
               Left            =   2640
               TabIndex        =   8
               Tag             =   "X"
               Top             =   1680
               Width           =   5055
            End
            Begin VB.TextBox txtDescripcion 
               Height          =   360
               Left            =   2640
               TabIndex        =   7
               Tag             =   "X"
               Top             =   720
               Width           =   7455
            End
            Begin VB.TextBox txtPrecioBase 
               Height          =   360
               Left            =   2640
               MaxLength       =   6
               TabIndex        =   6
               Tag             =   "X"
               Top             =   1200
               Width           =   2175
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Código Alternativo:"
               Height          =   240
               Left            =   600
               TabIndex        =   16
               Top             =   300
               Width           =   1905
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción:"
               Height          =   240
               Left            =   1305
               TabIndex        =   15
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Stock:"
               Height          =   240
               Left            =   1860
               TabIndex        =   14
               Top             =   1740
               Width           =   660
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Precio Base:"
               Height          =   240
               Left            =   1290
               TabIndex        =   13
               Top             =   1260
               Width           =   1230
            End
            Begin VB.Label lblActivo 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "1"
               Height          =   360
               Left            =   2640
               TabIndex        =   12
               Tag             =   "X"
               Top             =   2160
               Width           =   1695
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Activo:"
               Height          =   240
               Left            =   1800
               TabIndex        =   11
               Top             =   2240
               Width           =   720
            End
            Begin VB.Label lblIdProducto 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   6600
               TabIndex        =   10
               Tag             =   "X"
               Top             =   240
               Visible         =   0   'False
               Width           =   1695
            End
         End
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   1
         Top             =   480
         Width           =   9975
      End
      Begin MSComctlLib.ListView lvProducto 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   2
         Top             =   960
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   11456
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Busqueda"
         Height          =   240
         Left            =   -74640
         TabIndex        =   3
         Top             =   547
         Width           =   945
      End
   End
   Begin ToolBar.McToolBar mtbProducto 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
      ButtonIcon1     =   "frmMantVDProducto.frx":1B90
      ButtonToolTipIcon1=   1
      ButtonIconAllignment1=   0
      ButtonCaption2  =   "&Guardar"
      ButtonIcon2     =   "frmMantVDProducto.frx":286A
      ButtonToolTipIcon2=   1
      ButtonIconAllignment2=   0
      ButtonCaption3  =   "&Modificar"
      ButtonIcon3     =   "frmMantVDProducto.frx":3544
      ButtonToolTipIcon3=   1
      ButtonIconAllignment3=   0
      ButtonCaption4  =   "&Cancelar"
      ButtonIcon4     =   "frmMantVDProducto.frx":421E
      ButtonToolTipIcon4=   1
      ButtonIconAllignment4=   0
      ButtonCaption5  =   "&Desactivar"
      ButtonIcon5     =   "frmMantVDProducto.frx":4EF8
      ButtonToolTipIcon5=   1
      ButtonIconAllignment5=   0
      ButtonCaption6  =   "&Activar"
      ButtonIcon6     =   "frmMantVDProducto.frx":5BD2
      ButtonToolTipIcon6=   1
      ButtonIconAllignment6=   0
      ButtonCaption7  =   "&Eliminar"
      ButtonIcon7     =   "frmMantVDProducto.frx":68AC
      ButtonToolTipIcon7=   1
      ButtonIconAllignment7=   0
   End
End
Attribute VB_Name = "frmMantVDProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private pIDempresa As Integer
Private oRSPrecios As New ADODB.Recordset

Private Function ObtenerMaximoIDPrecio(rs As ADODB.Recordset) As Integer

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

            If Not IsNull(rs!idprecio.Value) Then
                If rs!idprecio.Value > maxID Then
                    maxID = rs!idprecio.Value

                End If

            End If

            rs.MoveNext
        Loop
    
        ' Restaurar posición original
        rs.bookmark = bookmark

    End If

    Dim itemx      As Object

    Dim IdPreMaxLV As Integer

    IdPreMaxLV = 0

    For Each itemx In Me.lvPrecios.ListItems

        If itemx.Tag > IdPreMaxLV Then
            IdPreMaxLV = itemx.Tag

        End If

    Next

    If IdPreMaxLV > maxID Then
        ObtenerMaximoIDPrecio = IdPreMaxLV + 1
    Else
        ObtenerMaximoIDPrecio = maxID + 1

    End If

    Exit Function
    
ErrorHandler:
    MsgBox "Error al buscar máximo ID: " & Err.Description, vbExclamation
    ObtenerMaximoIDPrecio = -1 ' Valor de error

End Function

Private Function validaCategorias() As Boolean
    On Error GoTo xValida
    
    MousePointer = vbHourglass
    
    'Configurar y ejecutar el comando
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_LIST_CATEGORIA]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Prepared = True
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
     
    Set oRSmain = oCmdEjec.Execute
    
    'Validar categorías
    Dim itemx As Object
    Dim bCategoriaValida As Boolean
    bCategoriaValida = False
    
    Do While Not oRSmain.EOF
        If oRSmain!ide <> -1 Then
            bCategoriaValida = False 'Resetear para cada categoría
            
            'Buscar en el ListView
            For Each itemx In Me.lvPrecios.ListItems
                If itemx.SubItems(1) = CStr(oRSmain!ide) And itemx.SubItems(3) = "SI" Then
                    bCategoriaValida = True
                    Exit For
                End If
            Next
            
            'Si no se encontró una categoría válida, salir inmediatamente
            If Not bCategoriaValida Then
                Exit Do
            End If
        End If
        oRSmain.MoveNext
    Loop
    
    CerrarConexion True
    MousePointer = vbDefault
    
    'Retornar resultado
    validaCategorias = bCategoriaValida
    Exit Function
    
xValida:
CerrarConexion True
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo
    validaCategorias = False
End Function


Private Sub cargarCategorias()

    On Error GoTo xCarga

    MousePointer = vbHourglass
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_LIST_CATEGORIA]"
    oCmdEjec.Prepared = True
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
     
    Set oRSmain = oCmdEjec.Execute
     
    If Not oRSmain.EOF Then

        ' Crear Recordsets temporales en memoria
        Dim orsTEMP1 As New ADODB.Recordset

        orsTEMP1.CursorLocation = adUseClient
        orsTEMP1.Fields.Append oRSmain(0).Name, oRSmain(0).Type, oRSmain(0).DefinedSize
        orsTEMP1.Fields.Append oRSmain(1).Name, oRSmain(1).Type, oRSmain(1).DefinedSize
        orsTEMP1.Open
    
        ' Copiar datos del Recordset
        oRSmain.MoveFirst

        Do Until oRSmain.EOF
            orsTEMP1.AddNew
            orsTEMP1.Fields(0).Value = oRSmain.Fields(0).Value
            orsTEMP1.Fields(1).Value = oRSmain.Fields(1).Value
            orsTEMP1.Update
            oRSmain.MoveNext
        Loop
        
        ' Configurar DatAddCategoria
        Set Me.datAddCategoria.RowSource = orsTEMP1
        Me.datAddCategoria.ListField = orsTEMP1.Fields(1).Name
        Me.datAddCategoria.BoundColumn = orsTEMP1.Fields(0).Name
        Me.datAddCategoria.BoundText = -1
        
        CerrarConexion True
        MousePointer = vbDefault

    End If
     
    Exit Sub
xCarga:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub


Sub Mandar_Datos()
    MousePointer = vbHourglass
EliminarRegistrosRecordSet oRSPrecios
    With Me.lvProducto
        Me.lblIdProducto.Caption = .SelectedItem.Tag
        Me.txtCodAlternativo.Text = .SelectedItem.Text
        Me.txtDescripcion.Text = .SelectedItem.SubItems(1)
        Me.lblActivo.Caption = .SelectedItem.SubItems(2)
        
        LimpiaParametros oCmdEjec, True
        oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_DATOS_ADICIONALES]"
        oCmdEjec.CommandType = adCmdStoredProc
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adInteger, adParamInput, , Me.lblIdProducto.Caption)
        oCmdEjec.Execute
        
        Set oRSmain = oCmdEjec.Execute
        
        If Not oRSmain.EOF Then
        Me.txtPrecioBase.Text = oRSmain!PRE
        Me.txtStock.Text = oRSmain!stk
        Me.lvPrecios.ListItems.Clear
        Dim ORSt As ADODB.Recordset
        Set ORSt = oRSmain.NextRecordset
        Dim itemx As Object
        Do While Not ORSt.EOF
            Set itemx = Me.lvPrecios.ListItems.Add(, , ORSt!cat, Me.ilProducto.ListImages(1).Key, Me.ilProducto.ListImages(1).Key)
            itemx.Tag = ORSt!idpre
            itemx.SubItems(1) = ORSt!IDcat
            itemx.SubItems(2) = ORSt!PRE
            itemx.SubItems(3) = ORSt!ACT
               If ORSt!ACT = "NO" Then
                Me.lvPrecios.ListItems(itemx.Index).ForeColor = vbRed
                Me.lvPrecios.ListItems(itemx.Index).ListSubItems(1).ForeColor = vbRed
                Me.lvPrecios.ListItems(itemx.Index).ListSubItems(2).ForeColor = vbRed
                Me.lvPrecios.ListItems(itemx.Index).ListSubItems(3).ForeColor = vbRed

            End If
            ORSt.MoveNext
        Loop
        
        Set ORSt = oRSmain.NextRecordset
        Do While Not ORSt.EOF
        oRSPrecios.AddNew
        oRSPrecios.Fields(0).Value = ORSt!idp
        oRSPrecios.Fields(1).Value = ORSt!idc
        oRSPrecios.Update
            ORSt.MoveNext
        Loop
        
        

        End If

        CerrarConexion True
        Estado_Botones AntesDeActualizar

    End With

    MousePointer = vbDefault

End Sub

Private Sub productoSearch(xdato As String)
    MousePointer = vbHourglass

    On Error GoTo xSearch

    Me.lvProducto.ListItems.Clear
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_SEARCH]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)

    If Len(Trim(xdato)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 100, xdato)
    
    Set oRSmain = oCmdEjec.Execute

    If Not oRSmain.EOF Then

        Dim itemx As Object

        Do While Not oRSmain.EOF
            Set itemx = Me.lvProducto.ListItems.Add(, , oRSmain!codalt, Me.ilProducto.ListImages(1).Key, Me.ilProducto.ListImages(1).Key)
            itemx.Tag = oRSmain!ide
            itemx.SubItems(1) = oRSmain!prod
            itemx.SubItems(2) = oRSmain!ACT

            If oRSmain!ACT = "NO" Then
                Me.lvProducto.ListItems(itemx.Index).ForeColor = vbRed
                Me.lvProducto.ListItems(itemx.Index).ListSubItems(1).ForeColor = vbRed
                Me.lvProducto.ListItems(itemx.Index).ListSubItems(2).ForeColor = vbRed

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
            Me.mtbProducto.Button_Index = 1
            Me.mtbProducto.ButtonEnabled = True
            Me.mtbProducto.Button_Index = 2
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 3
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 4
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 5
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 6
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 7
            Me.mtbProducto.ButtonEnabled = False
            Me.lvPrecios.Enabled = False
            Me.SSTProducto.tab = 0

        Case Nuevo, Editar
            Me.lblActivo.Caption = "SI"
            Me.mtbProducto.Button_Index = 1
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 2
            Me.mtbProducto.ButtonEnabled = True
            Me.mtbProducto.Button_Index = 3
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 4
            Me.mtbProducto.ButtonEnabled = True
            Me.mtbProducto.Button_Index = 5
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 6
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 7
            Me.mtbProducto.ButtonEnabled = False
            Me.lvProducto.Enabled = False
            Me.txtSearch.Enabled = False
            Me.lvPrecios.Enabled = True
            Me.SSTProducto.tab = 1

        Case buscar
            Me.mtbProducto.Button_Index = 1
            Me.mtbProducto.ButtonEnabled = True
            Me.mtbProducto.Button_Index = 2
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 3
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 4
            Me.mtbProducto.ButtonEnabled = False
            Me.SSTProducto.tab = 0

        Case AntesDeActualizar
            Me.mtbProducto.Button_Index = 1
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 2
            Me.mtbProducto.ButtonEnabled = False
            Me.mtbProducto.Button_Index = 3
            Me.mtbProducto.ButtonEnabled = True
            Me.mtbProducto.Button_Index = 4
            Me.mtbProducto.ButtonEnabled = True

            If Me.lblActivo.Caption = "SI" Then
                Me.mtbProducto.Button_Index = 5
                Me.mtbProducto.ButtonEnabled = True
                Me.mtbProducto.Button_Index = 6
                Me.mtbProducto.ButtonEnabled = False

            Else
                Me.mtbProducto.Button_Index = 5
                Me.mtbProducto.ButtonEnabled = False
                Me.mtbProducto.Button_Index = 6
                Me.mtbProducto.ButtonEnabled = True
            
            End If

            Me.lvPrecios.Enabled = False
            Me.mtbProducto.Button_Index = 7
            Me.mtbProducto.ButtonEnabled = True
            Me.SSTProducto.tab = 1

    End Select

End Sub

Private Sub ConfigurarLV()
Me.lvProducto.Icons = Me.ilProducto
Me.lvProducto.SmallIcons = Me.ilProducto

Me.lvPrecios.Icons = Me.ilProducto
Me.lvPrecios.SmallIcons = Me.ilProducto

With Me.lvProducto
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "CODIGO"
    .ColumnHeaders.Add , , "PRODUCTO", 5000
    .ColumnHeaders.Add , , "ACTIVO"
End With

With Me.lvPrecios
  .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "CATEGORIA", 3000
    .ColumnHeaders.Add , , "IDCAT", 0
    .ColumnHeaders.Add , , "PRECIO"
    .ColumnHeaders.Add , , "ACTIVO"

End With
End Sub

Private Sub cmdAdd_Click()

    If Me.datAddCategoria.BoundText = -1 Then
        MsgBox "Debe elegir la categoria.", vbInformation, Pub_Titulo
        Me.datAddCategoria.SetFocus
    ElseIf Len(Trim(Me.txtAddPrecio.Text)) = 0 Then
        MsgBox "Debe ingresar el Precio.", vbInformation, Pub_Titulo
        Me.txtAddPrecio.SetFocus
    ElseIf Me.txtAddPrecio.Text <= 0 Then
        MsgBox "Precio ingresado incorrecto.", vbInformation, Pub_Titulo
        Me.txtAddPrecio.SetFocus
        Me.txtAddPrecio.SelStart = 0
        Me.txtAddPrecio.SelLength = Len(Me.txtAddPrecio.Text)
  
    Else

        Dim itemx As Object

        Dim xData As Boolean, xIDpreMax As Integer

        xData = False

        If Me.lvPrecios.ListItems.count = 0 Then
           
            If oRSPrecios.RecordCount = 0 Then
                xIDpreMax = 1
            Else
                xIDpreMax = ObtenerMaximoIDPrecio(oRSPrecios)

            End If

            Set itemx = Me.lvPrecios.ListItems.Add(, , Me.datAddCategoria.Text, Me.ilProducto.ListImages(1).Key, Me.ilProducto.ListImages(1).Key)
            itemx.Tag = xIDpreMax
            itemx.SubItems(1) = Me.datAddCategoria.BoundText
            itemx.SubItems(2) = Me.txtAddPrecio.Text
            itemx.SubItems(3) = "SI"
        Else

            For Each itemx In Me.lvPrecios.ListItems

                If itemx.SubItems(1) = Me.datAddCategoria.BoundText And itemx.SubItems(3) = "SI" Then
                    xData = True
                    Exit For

                End If

            Next
        
            If xData Then
                MsgBox "Categoria ya se encuentra en lista.", vbInformation, Pub_Titulo
                Me.datAddCategoria.SetFocus
                Exit Sub
            Else
                xIDpreMax = ObtenerMaximoIDPrecio(oRSPrecios)
                Set itemx = Me.lvPrecios.ListItems.Add(, , Me.datAddCategoria.Text, Me.ilProducto.ListImages(1).Key, Me.ilProducto.ListImages(1).Key)
                itemx.Tag = xIDpreMax
                itemx.SubItems(1) = Me.datAddCategoria.BoundText
                itemx.SubItems(2) = Me.txtAddPrecio.Text
                itemx.SubItems(3) = "SI"

            End If

        End If
        
        Me.datAddCategoria.BoundText = -1
        Me.txtAddPrecio.Text = ""
        
        Me.datAddCategoria.SetFocus

    End If

End Sub

Private Sub cmdDel_Click()

    If Me.lvPrecios.ListItems.count = 0 Then Exit Sub
    If Me.lvPrecios.SelectedItem Is Nothing Then Exit Sub

    With Me.lvPrecios.SelectedItem
        oRSPrecios.AddNew
        oRSPrecios.Fields(0).Value = .Tag
        oRSPrecios.Fields(1).Value = .SubItems(1)
        oRSPrecios.Update

    End With
    
    Me.lvPrecios.ListItems.Remove Me.lvPrecios.SelectedItem.Index

End Sub

Private Sub datAddCategoria_KeyPress(KeyAscii As Integer)
HandleEnterKey KeyAscii, Me.txtAddPrecio
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ' Crear Recordsets temporales en memoria
If oRSPrecios.State = adStateOpen Then oRSPrecios.Close
    oRSPrecios.CursorLocation = adUseClient
    oRSPrecios.Fields.Append "idprecio", adInteger ' oRSmain(0).Name, oRSmain(0).Type, oRSmain(0).DefinedSize
    oRSPrecios.Fields.Append "idcategoria", adInteger ' oRSmain(1).Name, oRSmain(1).Type, oRSmain(1).DefinedSize
    oRSPrecios.Open
        
    pIDempresa = devuelveIDempresaXdefecto
    ConfigurarLV
    DesactivarControles Me
    Estado_Botones InicializarFormulario
    productoSearch Me.txtSearch.Text
    CentrarFormulario MDIForm1, Me
    
End Sub

Private Sub lvPrecios_DblClick()

    If Me.lvPrecios.ListItems.count = 0 Then Exit Sub
    frmMantVDProductoPrecio.gIDempresa = pIDempresa
    frmMantVDProductoPrecio.gIDcategoria = Me.lvPrecios.SelectedItem.SubItems(1)
    frmMantVDProductoPrecio.txtAddPrecio.Text = Me.lvPrecios.SelectedItem.SubItems(2)
frmMantVDProductoPrecio.lblCategoria.Caption = Me.lvPrecios.SelectedItem.Text
    frmMantVDProductoPrecio.ComActivo.ListIndex = IIf(Me.lvPrecios.SelectedItem.SubItems(3) = "NO", 0, 1)
    frmMantVDProductoPrecio.Show vbModal

End Sub

Private Sub lvProducto_DblClick()
Mandar_Datos
End Sub



Private Sub mtbProducto_Click(ByVal ButtonIndex As Long)

    Select Case ButtonIndex

        Case 1
            EliminarRegistrosRecordSet oRSPrecios
            ActivarControles Me
            LimpiarControles Me
            Me.lvPrecios.ListItems.Clear
            cargarCategorias
            Estado_Botones Nuevo
            Me.txtCodAlternativo.SetFocus
            VNuevo = True

        Case 2

            If Len(Trim(Me.txtCodAlternativo.Text)) = 0 Then
                MsgBox "Debe ingresar el [Código Alternativo]", vbCritical, Pub_Titulo
                Me.txtCodAlternativo.SetFocus
            ElseIf Len(Trim(Me.txtDescripcion.Text)) = 0 Then
                MsgBox "Debe ingresar la Descripción del Producto.", vbCritical, Pub_Titulo
                Me.txtDescripcion.SetFocus
            ElseIf Len(Trim(Me.txtPrecioBase.Text)) = 0 And Me.lvPrecios.ListItems.count = 0 Then
                MsgBox "Debe agregar un [Precio Base] si no va a registrar precios por Caracteristicas.", vbCritical, Pub_Titulo
                Me.datAddCategoria.SetFocus
            ElseIf val(Trim(Me.txtPrecioBase.Text)) <= 0 And Me.lvPrecios.ListItems.count = 0 Then
                MsgBox "[Precio Base] incorrecto,Debe ser mayor a Cero (0)" + vbCrLf + "si no va a registrar precios por Caracteristicas.", vbCritical, Pub_Titulo
                Me.datAddCategoria.SetFocus
            ElseIf Len(Trim(Me.txtStock.Text)) = 0 Then
                MsgBox "Debe ingresar el [Stock] del Producto.", vbCritical, Pub_Titulo
                Me.txtStock.SetFocus
            ElseIf Not validaCategorias Then
                MsgBox "Faltan Categorias que debe agregar.", vbCritical, Pub_Titulo
            Else
                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True

                If VNuevo Then
                    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_REGISTER]"
                Else
                    oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_UPDATE]"

                End If

                On Error GoTo grabar

                Dim Smensaje As String

                Dim vIDz     As Integer

                Smensaje = ""
                vIDz = 0

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, 2, pIDempresa)

                If Not VNuevo Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adInteger, adParamInput, , Me.lblIdProducto.Caption)
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODALTERNO", adVarChar, adParamInput, 20, Trim(Me.txtCodAlternativo.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCRIPCION", adVarChar, adParamInput, 100, Trim(Me.txtDescripcion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRECIO", adDouble, adParamInput, , Trim(Me.txtPrecioBase.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STOCK", adInteger, adParamInput, , Trim(Me.txtStock.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@UDUSARIO", adVarChar, adParamInput, 20, LK_CODUSU)

                Dim strItems As String

                strItems = ""

                Dim f As Integer
    
                If Me.lvPrecios.ListItems.count <> 0 Then
                    strItems = "<r>"

                    For f = 1 To Me.lvPrecios.ListItems.count
                        strItems = strItems & "<d "
                        strItems = strItems & "idp=""" & Me.lvPrecios.ListItems(f).Tag & """ "
                        strItems = strItems & "idc=""" & Me.lvPrecios.ListItems(f).SubItems(1) & """ "
                        strItems = strItems & "pr=""" & Me.lvPrecios.ListItems(f).SubItems(2) & """ "
                        strItems = strItems & "st=""" & IIf(Me.lvPrecios.ListItems(f).SubItems(3) = "NO", "0", "1") & """ "
                        strItems = strItems & "/>"
                    Next
                    strItems = strItems & "</r>"

                End If

                If Len(Trim(strItems)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@XPRECIOS", adVarChar, adParamInput, 4000, strItems)
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        DesactivarControles Me
                        Estado_Botones grabar
                        Me.lvProducto.Enabled = True
                        Me.txtSearch.Enabled = True
                        CerrarConexion True
                        productoSearch Me.txtSearch.Text
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

        Case 3
            VNuevo = False
            cargarCategorias
            Estado_Botones Editar
            ActivarControles Me
            Me.txtSearch.Enabled = False

        Case 4
            EliminarRegistrosRecordSet oRSPrecios
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvPrecios.Enabled = False
            Me.lvProducto.Enabled = True
            Me.txtSearch.Enabled = True
            Me.txtSearch.SetFocus

        Case 5

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            
                On Error GoTo Desactiva

                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_STATE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.lblIdProducto.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STATE", adBoolean, adParamInput, , False)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        LimpiarControles Me
                        Estado_Botones Desactivar
                        Me.lvPrecios.ListItems.Clear
                        Me.lvProducto.Enabled = True
                        productoSearch Me.txtSearch.Text
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
                oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_STATE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.lblIdProducto.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STATE", adBoolean, adParamInput, , True)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        LimpiarControles Me
                        Estado_Botones Activar
                        Me.lvPrecios.ListItems.Clear
                        Me.lvProducto.Enabled = True
                        productoSearch Me.txtSearch.Text
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
                oCmdEjec.CommandText = "[vd].[USP_PRODUCTO_DELETE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.lblIdProducto.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
              
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        DesactivarControles Me
                        Me.lvPrecios.ListItems.Clear
                        Estado_Botones Eliminar
                        Me.lvProducto.Enabled = True
                        Me.txtSearch.Enabled = True
                
                        productoSearch Me.txtSearch.Text
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

Private Sub txtAddPrecio_Change()
ValidarSoloNumerosPunto Me.txtAddPrecio
End Sub

Private Sub txtAddPrecio_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumerosPunto(Me.txtAddPrecio, KeyAscii)
 HandleEnterKey KeyAscii, Me.cmdAdd
End Sub

Private Sub txtPrecioBase_Change()
ValidarSoloNumerosPunto Me.txtPrecioBase
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
HandleEnterKey KeyAscii, Me.datAddCategoria
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
HandleEnterKey KeyAscii, Me.txtPrecioBase
End Sub

Private Sub txtCodAlternativo_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
HandleEnterKey KeyAscii, Me.txtDescripcion
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then productoSearch Me.txtSearch.Text
End Sub

Private Sub txtPrecioBase_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumerosPunto(Me.txtPrecioBase, KeyAscii)
HandleEnterKey KeyAscii, Me.txtStock
End Sub
