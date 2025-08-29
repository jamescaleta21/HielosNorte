VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{FEC367D0-B73E-4DD0-80FD-1F56BC27B04A}#1.0#0"; "McToolBar.ocx"
Begin VB.Form frmMantVDCategoria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Categorias [Venta Directa]"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMantVDCategoria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9450
   Begin ToolBar.McToolBar mtbCategoria 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
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
      ButtonIcon1     =   "frmMantVDCategoria.frx":058A
      ButtonToolTipIcon1=   1
      ButtonIconAllignment1=   0
      ButtonCaption2  =   "&Guardar"
      ButtonIcon2     =   "frmMantVDCategoria.frx":1264
      ButtonToolTipIcon2=   1
      ButtonIconAllignment2=   0
      ButtonCaption3  =   "&Modificar"
      ButtonIcon3     =   "frmMantVDCategoria.frx":1F3E
      ButtonToolTipIcon3=   1
      ButtonIconAllignment3=   0
      ButtonCaption4  =   "&Cancelar"
      ButtonIcon4     =   "frmMantVDCategoria.frx":2C18
      ButtonToolTipIcon4=   1
      ButtonIconAllignment4=   0
      ButtonCaption5  =   "&Desactivar"
      ButtonIcon5     =   "frmMantVDCategoria.frx":38F2
      ButtonToolTipIcon5=   1
      ButtonIconAllignment5=   0
      ButtonCaption6  =   "&Activar"
      ButtonIcon6     =   "frmMantVDCategoria.frx":45CC
      ButtonToolTipIcon6=   1
      ButtonIconAllignment6=   0
      ButtonCaption7  =   "&Eliminar"
      ButtonIcon7     =   "frmMantVDCategoria.frx":52A6
      ButtonToolTipIcon7=   1
      ButtonIconAllignment7=   0
   End
   Begin MSComctlLib.ImageList ilCategoria 
      Left            =   10680
      Top             =   3240
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
            Picture         =   "frmMantVDCategoria.frx":5F80
            Key             =   "category"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTCategoria 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8916
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
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmMantVDCategoria.frx":651A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvCategoria"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSearch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Categoria"
      TabPicture(1)   =   "frmMantVDCategoria.frx":6536
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "lblidCategoria"
      Tab(1).Control(4)=   "lblActivo"
      Tab(1).Control(5)=   "txtDenominacion"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtDenominacion 
         Height          =   360
         Left            =   -71160
         TabIndex        =   3
         Tag             =   "X"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtSearch 
         Height          =   360
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   7695
      End
      Begin MSComctlLib.ListView lvCategoria 
         Height          =   4095
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
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
      Begin VB.Label lblActivo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   -71160
         TabIndex        =   9
         Tag             =   "X"
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lblidCategoria 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   -71160
         TabIndex        =   8
         Tag             =   "X"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   240
         Left            =   -72000
         TabIndex        =   7
         Top             =   3060
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
         Height          =   240
         Left            =   -72720
         TabIndex        =   6
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id Categoria:"
         Height          =   240
         Left            =   -72555
         TabIndex        =   5
         Top             =   1500
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Busqueda"
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   540
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmMantVDCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private pIDempresa As Integer

Sub Mandar_Datos()

    With Me.lvCategoria
        Me.lblidCategoria.Caption = .SelectedItem.Text
        Me.txtDenominacion.Text = .SelectedItem.SubItems(1)
        Me.lblActivo.Caption = .SelectedItem.SubItems(2)
   
        Estado_Botones AntesDeActualizar

    End With

End Sub

Private Sub categoriaSearch(xdato As String)

    On Error GoTo xSearch

    Me.lvCategoria.ListItems.Clear
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_CATEGORIA_SEARCH]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)

    If Len(Trim(xdato)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 100, xdato)
    
    Set oRSmain = oCmdEjec.Execute
    MousePointer = vbHourglass

    If Not oRSmain.EOF Then

        Dim itemx As Object

        Do While Not oRSmain.EOF
            Set itemx = Me.lvCategoria.ListItems.Add(, , oRSmain!IDCATEGORIA, Me.ilCategoria.ListImages(1).Key, Me.ilCategoria.ListImages(1).Key)
            itemx.SubItems(1) = oRSmain!DENOMINACION
            itemx.SubItems(2) = oRSmain!activo

            If oRSmain!activo = "NO" Then
                Me.lvCategoria.ListItems(itemx.Index).ForeColor = vbRed
                Me.lvCategoria.ListItems(itemx.Index).ListSubItems(1).ForeColor = vbRed
                Me.lvCategoria.ListItems(itemx.Index).ListSubItems(2).ForeColor = vbRed

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
              Me.mtbCategoria.Button_Index = 1
            Me.mtbCategoria.ButtonEnabled = True
            Me.mtbCategoria.Button_Index = 2
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 3
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 4
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 5
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 6
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 7
            Me.mtbCategoria.ButtonEnabled = False
            Me.SSTCategoria.Tab = 0

        Case Nuevo, Editar
            Me.lblActivo.Caption = "SI"
             Me.mtbCategoria.Button_Index = 1
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 2
            Me.mtbCategoria.ButtonEnabled = True
            Me.mtbCategoria.Button_Index = 3
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 4
            Me.mtbCategoria.ButtonEnabled = True
            Me.mtbCategoria.Button_Index = 5
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 6
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 7
            Me.mtbCategoria.ButtonEnabled = False
            Me.lvCategoria.Enabled = False
            Me.txtSearch.Enabled = False
            Me.SSTCategoria.Tab = 1

        Case buscar
            Me.mtbCategoria.Button_Index = 1
            Me.mtbCategoria.ButtonEnabled = True
            Me.mtbCategoria.Button_Index = 2
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 3
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 4
            Me.mtbCategoria.ButtonEnabled = False
            Me.SSTCategoria.Tab = 0

        Case AntesDeActualizar
                 Me.mtbCategoria.Button_Index = 1
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 2
            Me.mtbCategoria.ButtonEnabled = False
            Me.mtbCategoria.Button_Index = 3
            Me.mtbCategoria.ButtonEnabled = True
            Me.mtbCategoria.Button_Index = 4
            Me.mtbCategoria.ButtonEnabled = True

            If Me.lblActivo.Caption = "SI" Then
                Me.mtbCategoria.Button_Index = 5
                Me.mtbCategoria.ButtonEnabled = True
                Me.mtbCategoria.Button_Index = 6
                Me.mtbCategoria.ButtonEnabled = False

            Else
                Me.mtbCategoria.Button_Index = 5
                Me.mtbCategoria.ButtonEnabled = False
                Me.mtbCategoria.Button_Index = 6
                Me.mtbCategoria.ButtonEnabled = True
            
            End If
             Me.mtbCategoria.Button_Index = 7
            Me.mtbCategoria.ButtonEnabled = True
            Me.SSTCategoria.Tab = 1

    End Select

End Sub

Private Sub ConfigurarLV()
Me.lvCategoria.Icons = Me.ilCategoria
Me.lvCategoria.SmallIcons = Me.ilCategoria

With Me.lvCategoria
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "IDE"
    .ColumnHeaders.Add , , "CATEGORIA", 3000
    .ColumnHeaders.Add , , "ACTIVO"
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
'pIDempresa = devuelveIDempresaXdefecto
'ConfigurarLV
'DesactivarControles Me
'Estado_Botones InicializarFormulario
'categoriaSearch Me.txtSearch.Text
'CentrarFormulario MDIForm1, Me
End Sub

Private Sub lvCategoria_DblClick()
Mandar_Datos
End Sub

Private Sub mtbCategoria_Click(ByVal ButtonIndex As Long)

    Select Case ButtonIndex

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            VNuevo = True
            Me.txtDenominacion.SetFocus

        Case 2 'Guardar

            If Len(Trim(Me.txtDenominacion.Text)) = 0 Then
                MsgBox "Debe ingresar la Denominación", vbCritical, Pub_Titulo
                Me.txtDenominacion.SetFocus
          
            Else
                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True

                If VNuevo Then
                    oCmdEjec.CommandText = "[vd].[USP_CATEGORIA_REGISTER]"
                Else
                    oCmdEjec.CommandText = "[vd].[USP_CATEGORIA_UPDATE]"

                End If

                On Error GoTo grabar

                Dim Smensaje As String

                Dim vIDz     As Integer

                Smensaje = ""
                vIDz = 0

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, 2, pIDempresa)

                If Not VNuevo Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCATEGORIA", adInteger, adParamInput, , Me.lblidCategoria.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DENOMINACION", adVarChar, adParamInput, 100, Trim(Me.txtDenominacion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@UDUSARIO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        DesactivarControles Me
                        Estado_Botones grabar
                        Me.lvCategoria.Enabled = True
                        Me.txtSearch.Enabled = True
                        CerrarConexion True
                        categoriaSearch Me.txtSearch.Text
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
            Me.txtDenominacion.SetFocus
            Me.txtSearch.Enabled = False

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvCategoria.Enabled = True
            Me.txtSearch.Enabled = True
            Me.txtSearch.SetFocus
            
        Case 5 'Desactivar
            
            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            
                On Error GoTo Desactiva

                MousePointer = vbHourglass
                LimpiaParametros oCmdEjec, True
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[vd].[USP_CATEGORIA_STATE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCATEGORIA", adInteger, adParamInput, , Me.lblidCategoria.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STATE", adBoolean, adParamInput, , False)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        LimpiarControles Me
                        Estado_Botones Desactivar
                        Me.lvCategoria.Enabled = True
                        categoriaSearch Me.txtSearch.Text
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
                oCmdEjec.CommandText = "[vd].[USP_CATEGORIA_STATE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCATEGORIA", adInteger, adParamInput, , Me.lblidCategoria.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@STATE", adBoolean, adParamInput, , True)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        LimpiarControles Me
                        Estado_Botones Activar
                        Me.lvCategoria.Enabled = True
                        categoriaSearch Me.txtSearch.Text
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
                oCmdEjec.CommandText = "[vd].[USP_CATEGORIA_DELETE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCATEGORIA", adInteger, adParamInput, , Me.lblidCategoria.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
              
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        CerrarConexion True
                        DesactivarControles Me
                        Estado_Botones Eliminar
                        Me.lvCategoria.Enabled = True
                        Me.txtSearch.Enabled = True
                
                        categoriaSearch Me.txtSearch.Text
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

Private Sub tbCategoria_ButtonClick(ByVal Button As MSComctlLib.Button)


End Sub

Private Sub txtDenominacion_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then categoriaSearch Me.txtSearch.Text
End Sub
