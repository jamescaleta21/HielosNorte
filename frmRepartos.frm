VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmRepartos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar Vendedores a Repartidores"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17310
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   17310
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17310
      _ExtentX        =   30533
      _ExtentY        =   1164
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mostrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTReparto 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Reparto"
      TabPicture(0)   =   "frmRepartos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtObs"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmRepartos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "cmdCloud"
      Tab(1).Control(2)=   "dtpFechaFiltro"
      Tab(1).Control(3)=   "lvListado"
      Tab(1).Control(4)=   "cmdBuscar"
      Tab(1).Control(5)=   "cmdDelPedido"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdDelPedido 
         Caption         =   "Del"
         Height          =   360
         Left            =   -58920
         TabIndex        =   25
         Top             =   1440
         Width           =   750
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   360
         Left            =   -71280
         TabIndex        =   22
         Top             =   480
         Width           =   990
      End
      Begin MSComctlLib.ListView lvListado 
         Height          =   7455
         Left            =   -74760
         TabIndex        =   19
         Top             =   1080
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   13150
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
      Begin MSComCtl2.DTPicker dtpFechaFiltro 
         Height          =   315
         Left            =   -73920
         TabIndex        =   18
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   203358209
         CurrentDate     =   45614
      End
      Begin VB.TextBox txtObs 
         Height          =   975
         Left            =   1680
         TabIndex        =   14
         Top             =   7680
         Width           =   14535
      End
      Begin VB.Frame Frame2 
         Height          =   6975
         Left            =   7320
         TabIndex        =   9
         Top             =   480
         Width           =   9615
         Begin VB.CheckBox chkAll 
            Caption         =   "Marcar todos"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   6600
            Width           =   1935
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "Del"
            Height          =   360
            Left            =   8790
            TabIndex        =   15
            Top             =   840
            Width           =   750
         End
         Begin MSComctlLib.ListView lvData 
            Height          =   5775
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   10186
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capacidad Vehiculo:"
            Height          =   195
            Left            =   5565
            TabIndex        =   24
            Top             =   6540
            Width           =   1755
         End
         Begin VB.Label lblCapacidadVehiculo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7320
            TabIndex        =   23
            Top             =   6480
            Width           =   1395
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso:"
            Height          =   195
            Left            =   6840
            TabIndex        =   21
            Top             =   6180
            Width           =   480
         End
         Begin VB.Label lblPeso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7320
            TabIndex        =   20
            Top             =   6120
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6975
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   6855
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   480
            Left            =   120
            TabIndex        =   29
            Top             =   1440
            Width           =   5895
         End
         Begin VB.ComboBox ComRuta 
            Height          =   315
            ItemData        =   "frmRepartos.frx":0038
            Left            =   1800
            List            =   "frmRepartos.frx":0060
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1080
            Width           =   4095
         End
         Begin VB.CheckBox chkTodos 
            Caption         =   "Marcar todos"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   6600
            Width           =   1935
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   480
            Left            =   6020
            TabIndex        =   12
            Top             =   2520
            Width           =   750
         End
         Begin MSDataListLib.DataCombo DatRepartidor 
            Height          =   315
            Left            =   1800
            TabIndex        =   4
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DatVendedor 
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvPedidos 
            Height          =   4215
            Left            =   120
            TabIndex        =   3
            Top             =   2280
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   7435
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ruta:"
            Height          =   195
            Left            =   1110
            TabIndex        =   27
            Top             =   1140
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repartidor:"
            Height          =   195
            Left            =   600
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   675
            TabIndex        =   7
            Top             =   780
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pedidos:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdCloud 
         Caption         =   "Cerrar Repartos"
         Height          =   480
         Left            =   -65640
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   17
         Top             =   660
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   7680
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRepartos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private orsRepartidor As ADODB.Recordset
Private Sub chkAll_Click()
If Me.lvData.ListItems.count = 0 Then Exit Sub
    Dim lvItem As Object

    If Me.chkAll.Value Then
        Me.chkAll.Caption = "Desmarcar todos"

        For Each lvItem In Me.lvData.ListItems
            lvItem.Checked = True
        Next
    Else
        Me.chkAll.Caption = "Marcar todos"
        For Each lvItem In Me.lvData.ListItems
            lvItem.Checked = False
        Next
    End If
End Sub

Private Sub chkTodos_Click()
If Me.lvPedidos.ListItems.count = 0 Then Exit Sub

    Dim lvItem As Object

    If Me.chkTodos.Value Then
        Me.chkTodos.Caption = "Desmarcar todos"

        For Each lvItem In Me.lvPedidos.ListItems
            lvItem.Checked = True
        Next
    Else
        Me.chkTodos.Caption = "Marcar todos"
        For Each lvItem In Me.lvPedidos.ListItems
            lvItem.Checked = False
        Next
    End If

End Sub

Private Sub cmdAdd_Click()
    Dim lvItem As Object, xFound As Boolean

    xFound = False

    For Each lvItem In Me.lvPedidos.ListItems
        If lvItem.Checked Then
            xFound = True
            Exit For
        End If

    Next

    If Not xFound Then
        MsgBox "No ha marcado ningun pedido.", vbInformation, Pub_Titulo
        Exit Sub
    End If

    Dim lvItemF As Object, xRepetidos As String, xFoundR As Boolean

    xFoundR = False

    For Each lvItem In Me.lvPedidos.ListItems
        If lvItem.Checked Then
            For Each lvItemF In Me.lvData.ListItems
                If lvItem.Text = lvItemF.Text Then
                    xRepetidos = xRepetidos + vbCrLf + lvItem.Text
                    xFoundR = True
                    Exit For
                End If
            Next
        
            If Not xFoundR Then
                Set itemX = Me.lvData.ListItems.Add(, , lvItem.Text)
                itemX.SubItems(1) = Me.DatVendedor.BoundText
                itemX.SubItems(2) = lvItem.SubItems(1)
                itemX.SubItems(3) = lvItem.SubItems(2)
            End If
       
            xFoundR = False

        End If

    Next

    If Len(Trim(xRepetidos)) <> 0 Then
        MsgBox "los siguientes Pedidos ya se encuentran en la lista." + vbCrLf + xRepetidos
    End If
    
     For Each lvItem In Me.lvPedidos.ListItems
        lvItem.Checked = False
    Next
    SumarPedidosAsignados
    Me.chkTodos.Caption = "Marcar todos"
    Me.chkTodos.Value = 0
End Sub

Private Sub cmdBuscar_Click()
cargarRepartos
End Sub

Private Sub cmdCloud_Click()
If MsgBox("¿Está seguro que desea cerrar los Repartos?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub

    LimpiaParametros oCmdEjec
    MousePointer = vbHourglass
    oCmdEjec.CommandText = "[dbo].[USP_REPARTO_SYNC]"
    oCmdEjec.CommandTimeout = 1000
     oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTO", adInteger, adParamInput, , Me.lvListado.SelectedItem.Tag)

    Dim orsExe As ADODB.Recordset

    Dim xDAtos As String

    Set orsExe = oCmdEjec.Execute
    MousePointer = vbDefault

    If Not orsExe.EOF Then
        xDAtos = orsExe.Fields(0).Value
        MsgBox Split(xDAtos, "=")(1), vbInformation, Pub_Titulo

    End If

End Sub

Private Sub cmdDel_Click()
 Dim lvItem As Object, xFound As Boolean

    xFound = False

    For Each lvItem In Me.lvData.ListItems

        If lvItem.Checked Then
            xFound = True
            Exit For

        End If

    Next

    If Not xFound Then
        MsgBox "No ha marcado ningun pedido.", vbInformation, Pub_Titulo
        Exit Sub

    End If
    
   Dim I As Integer

    ' Recorre los elementos desde el final hacia el principio
    ' para evitar problemas al eliminar elementos mientras se recorre la colección.
    For I = Me.lvData.ListItems.count To 1 Step -1
        ' Verifica si el elemento está seleccionado (Check = True)
        If Me.lvData.ListItems(I).Checked = True Then
            ' Elimina el elemento del ListView
            Me.lvData.ListItems.Remove I
        End If
    Next I
   SumarPedidosAsignados
End Sub

Private Sub cmdDelPedido_Click()

    If MsgBox("Desea continuar con el proceso.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    If Me.lvListado.ListItems.count = 0 Then
        MsgBox "No hay nada para eliminar.", vbCritical, Pub_Titulo
        Exit Sub

    End If

    If Me.lvListado.SelectedItem.SubItems(3) = "SI" Then
        MsgBox "No puede eliminar el Reparto cerrado.", vbInformation, Pub_Titulo
        Exit Sub

    End If
    
    Dim strMotivo As String

    strMotivo = InputBox("Ingrese el Motivo dee Eliminación:", Pub_Titulo)
    
    If Len(Trim(strMotivo)) = 0 Then Exit Sub

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_REPARTO_ELIMINAR]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTO", adInteger, adParamInput, , Me.lvListado.SelectedItem.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MOTIVO", adVarChar, adParamInput, 200, strMotivo)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)
    
    Dim oRSdelete As ADODB.Recordset

    Set oRSdelete = oCmdEjec.Execute
    
    If Not oRSdelete.EOF Then
        If Split(oRSdelete.Fields(0).Value, "=")(0) = "0" Then
            MsgBox Split(oRSdelete.Fields(0).Value, "=")(1), vbInformation, Pub_Titulo
            cargarRepartos
        Else
            MsgBox Split(oRSdelete.Fields(0).Value, "=")(1), vbCritical, Pub_Titulo
        End If
    End If

End Sub

Private Sub cmdSearch_Click()
    Me.lvPedidos.ListItems.Clear

    If Me.DatVendedor.BoundText <> -1 Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "[dbo].[USP_REPARTO_LOADPEDIDOS]"
        oCmdEjec.CommandType = adCmdStoredProc
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SUBRUTA", adInteger, adParamInput, , Me.ComRuta.ListIndex - 1)

        Dim ORSd As ADODB.Recordset

        Set ORSd = oCmdEjec.Execute
    
        Do While Not ORSd.EOF
            Set itemX = Me.lvPedidos.ListItems.Add(, , ORSd!idpedido)
            itemX.SubItems(1) = ORSd!Nombre
            itemX.SubItems(2) = ORSd!peso
            itemX.SubItems(3) = ORSd!subruta
            itemX.SubItems(4) = ORSd!obs
            ORSd.MoveNext
        Loop

    End If

End Sub

Private Sub DatRepartidor_Change()
Me.lvPedidos.ListItems.Clear
Me.lblPeso.Caption = "0.00"
Me.lblCapacidadVehiculo.Caption = "0.00"
Me.lvData.ListItems.Clear
If Me.DatRepartidor.BoundText <> -1 Then
    orsRepartidor.Filter = "cod=" & Me.DatRepartidor.BoundText
    If Not orsRepartidor.EOF Then
        Me.lblCapacidadVehiculo.Caption = orsRepartidor!Cap
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.dtpFechaFiltro.Value = LK_FECHA_DIA
ConfigurarLV
cargarRepartidoresVendedores
cargarRepartos
Desactivar
Me.ComRuta.ListIndex = 0
Me.Toolbar1.Buttons(2).Enabled = False
Me.Toolbar1.Buttons(3).Enabled = False
Me.Toolbar1.Buttons(4).Enabled = False
CenterMe Me
End Sub

Private Sub ConfigurarLV()

    With Me.lvPedidos
        .HideColumnHeaders = False
        .ColumnHeaders.Add , , "Nro Pedido", 1500
        .ColumnHeaders.Add , , "Cliente", 3000
        .ColumnHeaders.Add , , "Peso", 800
        .ColumnHeaders.Add , , "SubRuta", 1500
        .ColumnHeaders.Add , , "obs", 2500
        .FullRowSelect = True
        .View = lvwReport
        .CheckBoxes = True

    End With

    With Me.lvData
        .FullRowSelect = True
        .CheckBoxes = True
        .HideColumnHeaders = False
        .View = lvwReport
        .ColumnHeaders.Add , , "Nro Pedido", 1200
        .ColumnHeaders.Add , , "idVendedor", 0
        .ColumnHeaders.Add , , "Cliente", 3000
        .ColumnHeaders.Add , , "Peso", 1200

    End With

    With Me.lvListado
        .FullRowSelect = True
        .HideColumnHeaders = False
        .View = lvwReport
        .ColumnHeaders.Add , , "Reparto", 2500
        .ColumnHeaders.Add , , "Repartidor", 2500
        .ColumnHeaders.Add , , "Fecha"
        .ColumnHeaders.Add , , "Cerrado"
        .ColumnHeaders.Add , , "obs", 0
        .ColumnHeaders.Add , , "placa", 0

    End With

End Sub

Private Sub cargarRepartidoresVendedores()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_REPARTO_LOADUSUARIOS]"
oCmdEjec.CommandType = adCmdStoredProc
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

'Dim ors As ADODB.Recordset
Dim ORSt As ADODB.Recordset

Set orsRepartidor = oCmdEjec.Execute

Set Me.DatRepartidor.RowSource = orsRepartidor
Me.DatRepartidor.ListField = orsRepartidor(1).Name
Me.DatRepartidor.BoundColumn = orsRepartidor(0).Name
Me.DatRepartidor.BoundText = -1

Set ORSt = orsRepartidor.NextRecordset

Set Me.DatVendedor.RowSource = ORSt
Me.DatVendedor.ListField = ORSt(1).Name
Me.DatVendedor.BoundColumn = ORSt(0).Name
Me.DatVendedor.BoundText = -1


End Sub

Private Sub lvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.Toolbar1.Buttons(3).Enabled = True
End Sub

Private Sub lvPedidos_DblClick()
frmRepartos_Pedido.gIDpedido = Me.lvPedidos.SelectedItem.Text
frmRepartos_Pedido.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'Nuevo
            Me.SSTReparto.tab = 0
            Me.SSTReparto.TabEnabled(1) = False
            Me.Toolbar1.Buttons(1).Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = True
            Me.Toolbar1.Buttons(3).Enabled = False
            Me.Toolbar1.Buttons(4).Enabled = True
            limpiarPantalla
            Activar
            Me.Frame1.Enabled = True
            Me.Frame2.Enabled = True
            Me.txtObs.Enabled = True
            Me.DatRepartidor.SetFocus
        
        Case 2 'Grabar
        
            If Me.DatRepartidor.BoundText = -1 Then
                MsgBox "Debe elegir el repartidor.", vbCritical, Pub_Titulo
                Me.DatRepartidor.SetFocus
                Exit Sub

            End If
        
            If Me.DatVendedor.BoundText = -1 Then
                MsgBox "Debe elegir el Vendedor.", vbCritical, Pub_Titulo
                Me.DatVendedor.SetFocus
                Exit Sub

            End If
        
            If Me.lvData.ListItems.count = 0 Then
                MsgBox "Debe agregar pedidos para repartir.", vbCritical, Pub_Titulo
                Exit Sub

            End If
            
            If val(Me.lblPeso.Caption) > val(Me.lblCapacidadVehiculo.Caption) Then
                MsgBox "Carga superada del vehiculo.", vbInformation, Pub_Titulo
                Exit Sub

            End If

            Dim xDetalle As String

            xDetalle = ""

            If Me.lvData.ListItems.count = 0 Then
                MsgBox "No ha agregado algun pedido para proceder con el registro.", vbInformation, Pub_Titulo
                Exit Sub

            End If
         
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[dbo].[USP_REPARTO_REGISTER]"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTIDOR", adInteger, adParamInput, , Me.DatRepartidor.BoundText)
            
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PESO", adDouble, adParamInput, , Me.lblPeso.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)
            'AGREGANDO EL DETALLE
            
            If Me.lvData.ListItems.count <> 0 Then
                xDetalle = "["

                For Each lvItem In Me.lvData.ListItems

                    xDetalle = xDetalle & "{""" & "idp"":" & lvItem.Text & ",""idv"":" & lvItem.SubItems(1) & "},"
                Next
                xDetalle = Left(xDetalle, Len(xDetalle) - 1)
                xDetalle = xDetalle & "]"

            End If

            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DET", adLongVarWChar, adParamInput, Len(xDetalle), xDetalle)

            If Len(Trim(Me.txtObs.Text)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@OBS", adVarChar, adParamInput, 300, Trim(Me.txtObs.Text))

            Dim orsResult As ADODB.Recordset

            Set orsResult = oCmdEjec.Execute

            'AGREGANDO EL DETALLE
            Dim sMensaje() As String

            If Not orsResult.EOF Then
                sMensaje = Split(orsResult.Fields(0), "=")

                If sMensaje(0) = 0 Then

                    MousePointer = vbDefault
                    MsgBox sMensaje(1), vbInformation, Pub_Titulo

                Else
                    MousePointer = vbDefault
                    MsgBox sMensaje(1), vbCritical, Pub_Titulo

                End If

                limpiarPantalla
                Me.SSTReparto.TabEnabled(1) = True
                Me.Toolbar1.Buttons(1).Enabled = True
                Me.Toolbar1.Buttons(2).Enabled = False
                Me.Toolbar1.Buttons(3).Enabled = False
                Me.Toolbar1.Buttons(4).Enabled = False

            End If

        Case 3
            MousePointer = vbHourglass
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[dbo].[USP_REPARTO_REPORTE]"

            Dim orsDataResumen As ADODB.Recordset
        
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.lvListado.SelectedItem.SubItems(2))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTO", adInteger, adParamInput, , Me.lvListado.SelectedItem.Tag)
            
            Set orsDataResumen = oCmdEjec.Execute
            
            Dim orsDataMain As ADODB.Recordset

            Dim orsDataRes  As ADODB.Recordset
            
            Set orsDataMain = orsDataResumen.NextRecordset
            Set orsDataRes = orsDataMain.NextRecordset
            
            Set frmRepartos_View.pRSdata = orsDataMain
            Set frmRepartos_View.pRSdataRes = orsDataRes
            
            frmRepartos_View.pOBS = Me.lvListado.SelectedItem.SubItems(4)
            frmRepartos_View.pPLACA = Me.lvListado.SelectedItem.SubItems(5)
            frmRepartos_View.pIDREPARTO = Me.lvListado.SelectedItem.Tag
            frmRepartos_View.pREPARTIDOR = Me.lvListado.SelectedItem.SubItems(1)
            frmRepartos_View.pCantidad = orsDataResumen!cant
            frmRepartos_View.pVendedores = orsDataResumen!ListaVendedores
            frmRepartos_View.Caption = "Reparto Nro: " & Me.lvListado.SelectedItem.Tag
            MousePointer = vbDefault
            frmRepartos_View.Show vbModal

        Case 4
            Me.SSTReparto.TabEnabled(1) = True
            Me.Toolbar1.Buttons(1).Enabled = True
            Me.Toolbar1.Buttons(2).Enabled = False
            Me.Toolbar1.Buttons(3).Enabled = False
            Me.Toolbar1.Buttons(4).Enabled = False
            limpiarPantalla
            Desactivar

    End Select

End Sub

Private Sub limpiarPantalla()
Me.DatRepartidor.BoundText = -1
Me.DatVendedor.BoundText = -1
Me.txtObs.Text = ""
Me.lvData.ListItems.Clear
Me.lvPedidos.ListItems.Clear
Me.ComRuta.ListIndex = 0
End Sub

Private Sub cargarRepartos()
Me.lvListado.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_REPARTO_LIST]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.dtpFechaFiltro.Value)
    
    Dim ORSd As ADODB.Recordset
    Set ORSd = oCmdEjec.Execute
    
    Do While Not ORSd.EOF
        Set itemX = Me.lvListado.ListItems.Add(, , ORSd!reparto)
        itemX.Tag = ORSd!idreparto
        itemX.SubItems(1) = ORSd!repartidor
        itemX.SubItems(2) = ORSd!fecha
        itemX.SubItems(3) = ORSd!cloud
        itemX.SubItems(4) = ORSd!obs
        itemX.SubItems(5) = ORSd!placa
        ORSd.MoveNext
    Loop

End Sub

Private Sub SumarPedidosAsignados()
Dim lvItem As Object
Dim cPeso As Double
cPeso = 0

    For Each lvItem In Me.lvData.ListItems
       cPeso = cPeso + lvItem.SubItems(3)

    Next
    Me.lblPeso.Caption = cPeso
End Sub

Private Sub Desactivar()
Me.Frame1.Enabled = False
Me.Frame2.Enabled = False
Me.txtObs.Enabled = False
End Sub


Private Sub Activar()
Me.Frame1.Enabled = True
Me.Frame2.Enabled = True
Me.txtObs.Enabled = True
End Sub
