VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAsignarRutaClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar Clientes a Rutas de Vendedores"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16335
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
   ScaleHeight     =   9150
   ScaleWidth      =   16335
   Begin VB.Frame Frame2 
      Height          =   7935
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   16095
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   7560
         TabIndex        =   19
         Top             =   5160
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   720
         Left            =   7560
         Picture         =   "frmAsignarRutaClientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   990
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Del"
         Enabled         =   0   'False
         Height          =   720
         Left            =   7560
         Picture         =   "frmAsignarRutaClientes.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   990
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Up"
         Enabled         =   0   'False
         Height          =   720
         Left            =   7560
         Picture         =   "frmAsignarRutaClientes.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Width           =   990
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Down"
         Enabled         =   0   'False
         Height          =   720
         Left            =   7560
         Picture         =   "frmAsignarRutaClientes.frx":163E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   990
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   720
         Left            =   7560
         Picture         =   "frmAsignarRutaClientes.frx":1DA8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6840
         Width           =   990
      End
      Begin MSComctlLib.ListView lvDatos 
         Height          =   7335
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   12938
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvLibres 
         Height          =   7335
         Left            =   8640
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   12938
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
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   720
         Left            =   7560
         Picture         =   "frmAsignarRutaClientes.frx":2512
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6000
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Clientes Asignados"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   2475
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes no Asignados"
         Height          =   195
         Left            =   8640
         TabIndex        =   17
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label lblAsignados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6240
         TabIndex        =   16
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lblNoAsignados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   14820
         TabIndex        =   15
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   16095
      Begin VB.ComboBox ComDia 
         Height          =   315
         ItemData        =   "frmAsignarRutaClientes.frx":2C7C
         Left            =   8760
         List            =   "frmAsignarRutaClientes.frx":2C95
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   390
         Width           =   1935
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   600
         Left            =   11280
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DatVendedor 
         Height          =   315
         Left            =   2760
         TabIndex        =   0
         Top             =   390
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   1680
         TabIndex        =   13
         Top             =   450
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Día de la Semana:"
         Height          =   195
         Left            =   6960
         TabIndex        =   12
         Top             =   450
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmAsignarRutaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
'FECHA  :   21/12/2022
'AUTOR  :   JMENDOZA
'******************************** BITACORA DE CAMBIOS ********************************
'---------------------------------------------------------------
'CODIGO     |   NOMBRE      |   FECHA   |   MOTIVO
'---------------------------------------------------------------
'(@#)1-A        JMENDOZA        21/12/2022   Implementación de Secuencia del vendedor
'*************************************************************************************

Private Sub cmdAdd_Click()
Dim cant As String
cant = InputBox("ingrese Ubicación para el Cliente " & vbCrLf & Me.lvLibres.SelectedItem.Text)
If Len(Trim(cant)) = 0 Then
    MsgBox "No ha ingresado la Secuencia del Cliente.", vbCritical, Pub_Titulo
    Exit Sub
End If
If Not IsNumeric(cant) Then
    MsgBox "Secuencia ingresada incorrecta.", vbCritical, Pub_Titulo
    Exit Sub
End If
If Val(cant) > Me.lvDatos.ListItems.count + 1 Then
    MsgBox "Secuencia ingresada sobrepasa cantidad actual de clientes asignados.", vbInformation, Pub_Titulo
    Exit Sub
End If
If Val(cant) = 0 Then
    MsgBox "Secuencia ingresada incorrecta." + vbCrLf + "No puede ingresar cero (0).", vbInformation, Pub_Titulo
    Exit Sub
End If

Me.cmdAdd.Enabled = False
Dim itemx As Object

If Val(cant) = Me.lvDatos.ListItems.count + 1 Then
Set itemx = Me.lvDatos.ListItems.Add(, , cant)
Else
Set itemx = Me.lvDatos.ListItems.Add(cant, , cant)
End If

itemx.SubItems(1) = Me.lvLibres.SelectedItem.Text
itemx.SubItems(2) = Me.lvLibres.SelectedItem.SubItems(1)
itemx.Tag = Me.lvLibres.SelectedItem.Tag

Me.lvLibres.ListItems.Remove Me.lvLibres.SelectedItem.Index

ReordenarSecuencia


'''
'''Dim cant As Integer
'''If Me.lvLibres.SelectedItem.Tag <> 0 Then
'''    cant = InputBox("ingrese Ubicación")
'''    With lvDatos.ListItems.Add
'''        .Text = Me.lvLibres.SelectedItem.Text
'''        .SubItems(1) = Me.lvLibres.SelectedItem.SubItems(1)
'''        .Tag = Me.lvLibres.SelectedItem.Tag
'''       ' .SubItems(2) = cant
'''       ' .SubItems(3) = Me.lvDatos.SelectedItem.SubItems(2)
'''    End With
'''    Me.lvLibres.ListItems.Remove Me.lvLibres.SelectedItem.Index
'''End If
End Sub

Private Sub cmdBuscar_Click()
If Me.DatVendedor.BoundText = "" Then
    MsgBox "Debe elegir el vendedor.", vbInformation, Pub_Titulo
    Exit Sub
End If
If Me.ComDia.ListIndex = -1 Then
    MsgBox "Debe elegir el día.", vbCritical, Pub_Titulo
    Me.ComDia.SetFocus
    Exit Sub
End If
Me.lvDatos.ListItems.Clear
Me.lvLibres.ListItems.Clear
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_VENDEDOR_MUESTRARUTA]"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIA", adInteger, adParamInput, , Me.ComDia.ListIndex + 1)

Dim oRSdatos As ADODB.Recordset
Dim orsTemp As ADODB.Recordset

Set oRSdatos = oCmdEjec.Execute
Dim itemx As Object
Do While Not oRSdatos.EOF
    Set itemx = Me.lvDatos.ListItems.Add(, , oRSdatos!Sec)
    itemx.Tag = oRSdatos!IDe
    itemx.SubItems(1) = oRSdatos!nom
    itemx.SubItems(2) = oRSdatos!dir
    oRSdatos.MoveNext
Loop

Set orsTemp = oRSdatos.NextRecordset
Do While Not orsTemp.EOF
 Set itemx = Me.lvLibres.ListItems.Add(, , orsTemp!nom)
    itemx.Tag = orsTemp!IDe
    itemx.SubItems(1) = orsTemp!dir
    orsTemp.MoveNext
Loop
Me.cmdAdd.Enabled = False
Me.cmdDel.Enabled = False
Me.cmdUp.Enabled = False
Me.cmdDown.Enabled = False

Me.lblAsignados.Caption = Me.lvDatos.ListItems.count
Me.lblNoAsignados.Caption = Me.lvLibres.ListItems.count
End Sub

Private Sub cmdCancelar_Click()
Me.lvDatos.ListItems.Clear
Me.lvLibres.ListItems.Clear
Me.cmdAdd.Enabled = False
Me.cmdDel.Enabled = False
Me.cmdUp.Enabled = False
Me.cmdDown.Enabled = False
Me.lblAsignados.Caption = ""
Me.lblNoAsignados.Caption = ""
End Sub

Private Sub cmdDel_Click()

Me.cmdDel.Enabled = False

Dim itemx As Object
Set itemx = Me.lvLibres.ListItems.Add(, , Me.lvDatos.SelectedItem.SubItems(1))
itemx.SubItems(1) = Me.lvDatos.SelectedItem.SubItems(2)
itemx.Tag = Me.lvDatos.SelectedItem.Tag
Me.lvDatos.ListItems.Remove Me.lvDatos.SelectedItem.Index
ReordenarSecuencia
End Sub


Private Sub cmdDown_Click()
Dim itemx As Object
Dim itemSel As Integer
itemSel = Me.lvDatos.SelectedItem.Index + 2
If itemSel - 2 = Me.lvDatos.ListItems.count Then Exit Sub
If itemSel = 0 Then Exit Sub
Set itemx = Me.lvDatos.ListItems.Add(itemSel, , Me.lvDatos.SelectedItem.Text)
itemx.SubItems(1) = Me.lvDatos.SelectedItem.SubItems(1)
itemx.SubItems(2) = Me.lvDatos.SelectedItem.SubItems(2)
itemx.Tag = Me.lvDatos.SelectedItem.Tag
Me.lvDatos.ListItems.Remove Me.lvDatos.SelectedItem.Index
ReordenarSecuencia
Me.lvDatos.ListItems(itemSel - 1).Selected = True
End Sub

Private Sub cmdGrabar_Click()
If Me.lvDatos.ListItems.count = 0 Then
    MsgBox "Debe ingresar algun Cliente.", vbInformation, Pub_Titulo
Exit Sub
End If
On Error GoTo xSave
ReordenarSecuencia
Dim xASIGNADOS, xNOASIGNADOS As String
Dim I As Integer
For I = 1 To Me.lvDatos.ListItems.count
    xASIGNADOS = xASIGNADOS & Me.lvDatos.ListItems(I).Text & "," & Me.lvDatos.ListItems(I).Tag
    If I < Me.lvDatos.ListItems.count Then
        xASIGNADOS = xASIGNADOS + "|"
    End If
Next
xNOASIGNADOS = ""
For I = 1 To Me.lvLibres.ListItems.count
    xNOASIGNADOS = xNOASIGNADOS & Me.lvLibres.ListItems(I).Tag
    If I < Me.lvLibres.ListItems.count Then
        xNOASIGNADOS = xNOASIGNADOS + ","
    End If
Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_VENDEDORES_ASIGNA_CLIENTES]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIA", adInteger, adParamInput, , Me.ComDia.ListIndex + 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ASIGNADOS", adVarChar, adParamInput, -1, xASIGNADOS)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOASIGNADOS", adVarChar, adParamInput, -1, xNOASIGNADOS)
    oCmdEjec.Execute
    MsgBox "Datos procesados correctamente.", vbInformation, Pub_Titulo
Exit Sub
xSave:
MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub cmdImprimir_Click()
If Me.lvDatos.ListItems.count = 0 Then Exit Sub
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_VENDEDOR_MUESTRARUTA_PRINT]"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIA", adInteger, adParamInput, , Me.ComDia.ListIndex + 1)

Dim oRSdatos As ADODB.Recordset
Set oRSdatos = oCmdEjec.Execute
 
 Set frmAsignaRutaClientes_print.oRSdatos = oRSdatos
 frmAsignaRutaClientes_print.pVendedor = Me.DatVendedor.Text
 frmAsignaRutaClientes_print.pDia = Me.ComDia.Text
 frmAsignaRutaClientes_print.Show vbModal

End Sub

Private Sub cmdUp_Click()
Dim itemx As Object
Dim itemSel As Integer
itemSel = Me.lvDatos.SelectedItem.Index - 1
If itemSel = 0 Then Exit Sub
Set itemx = Me.lvDatos.ListItems.Add(itemSel, , Me.lvDatos.SelectedItem.Text)
itemx.SubItems(1) = Me.lvDatos.SelectedItem.SubItems(1)
itemx.SubItems(2) = Me.lvDatos.SelectedItem.SubItems(2)
itemx.Tag = Me.lvDatos.SelectedItem.Tag
Me.lvDatos.ListItems.Remove Me.lvDatos.SelectedItem.Index
ReordenarSecuencia
Me.lvDatos.ListItems(itemSel).Selected = True

End Sub


Private Sub ComDia_Click()
Me.lvDatos.ListItems.Clear
End Sub

Private Sub DatVendedor_Change()
Me.lvDatos.ListItems.Clear
End Sub

Private Sub Form_Load()
CargaVendedores
ConfiguraLvs
End Sub


Private Sub CargaVendedores()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "dbo.USP_VENDEDOR_LIST"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
Dim orsV As ADODB.Recordset
Set orsV = oCmdEjec.Execute

Set Me.DatVendedor.RowSource = orsV
Me.DatVendedor.BoundColumn = orsV.Fields(0).Name
Me.DatVendedor.ListField = orsV.Fields(1).Name

End Sub

Private Sub ConfiguraLvs()
With Me.lvDatos
    .ColumnHeaders.Add , , "#", 500
     .ColumnHeaders.Add , , "Cliente", 3000
    .ColumnHeaders.Add , , "Direccion", 3500
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = False
    .View = lvwReport
    .HideSelection = False
End With
With Me.lvLibres
     .ColumnHeaders.Add , , "Cliente", 3000
    .ColumnHeaders.Add , , "Direccion", 3500
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = False
    .View = lvwReport
    .HideSelection = False
End With
End Sub

Private Sub ReordenarSecuencia()

    Dim I As Integer

    For I = 1 To Me.lvDatos.ListItems.count
        Me.lvDatos.ListItems(I).Text = I
    Next

    Me.lblAsignados.Caption = Me.lvDatos.ListItems.count
    Me.lblNoAsignados.Caption = Me.lvLibres.ListItems.count

End Sub

Private Sub lvDatos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.cmdDel.Enabled = True
    Me.cmdUp.Enabled = True
    Me.cmdDown.Enabled = True

End Sub

Private Sub lvLibres_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdAdd.Enabled = True
End Sub
