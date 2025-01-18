VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmFormGenPedidos_Consulta 
   Caption         =   "Pedidos por Vendedor"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16560
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   16560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraReporte 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   16335
      Begin CRVIEWERLibCtl.CRViewer crvReporte 
         Height          =   7575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   16095
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   0   'False
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   -1  'True
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
   End
   Begin VB.Frame frafiltro 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   16335
      Begin VB.CheckBox chkres 
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12360
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         ItemData        =   "frmFormGenPedidos_Consulta.frx":0000
         Left            =   9840
         List            =   "frmFormGenPedidos_Consulta.frx":000D
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar"
         Height          =   600
         Left            =   14760
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   6120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   203358209
         CurrentDate     =   44900
      End
      Begin MSDataListLib.DataCombo DatVendedor 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   8880
         TabIndex        =   9
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   5400
         TabIndex        =   4
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   323
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmFormGenPedidos_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkres_Click()
If chkres.Value = 0 Then
chkres.Caption = "Detallado"
Else
chkres.Caption = "Resumido"
End If
End Sub

Private Sub cmdMostrar_Click()

    If Me.DatVendedor.BoundText = "" Then
        MsgBox "Debe elegir el vendedor.", vbInformation, Pub_Titulo
        Me.DatVendedor.SetFocus
        Exit Sub

    End If

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_PEDIDOS_VENDEDOR]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adChar, adParamInput, 8, FormatoFecha(Me.dtpFecha.Value))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FILTRO", adInteger, adParamInput, , Me.cboFiltro.ListIndex)

    Dim orsc As ADODB.Recordset

    Set orsc = oCmdEjec.Execute

    Dim vCAnt As Integer

    vCAnt = orsc!cant

    If chkres.Value = 0 Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_RESUMEN]"

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adChar, adParamInput, 8, FormatoFecha(Me.dtpFecha.Value))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FILTRO", adInteger, adParamInput, , Me.cboFiltro.ListIndex)

        Dim orsData As ADODB.Recordset

        Set orsData = oCmdEjec.Execute

        Dim vReporte   As CRAXDRT.Report

        Dim colParam   As CRAXDRT.ParameterFieldDefinitions

        Dim objParam   As CRAXDRT.ParameterFieldDefinition

        Dim objCrystal As New CRAXDRT.APPLICATION

        Set vReporte = objCrystal.OpenReport(PUB_RUTA_OTRO & "ReportePedidoVendedor.rpt")
    
        Set colParam = vReporte.ParameterFields
    
        For Each objParam In colParam

            Select Case objParam.ParameterFieldName

                Case "pCant"
                    objParam.AddCurrentValue CStr(vCAnt)

            End Select

        Next
    
        vReporte.Database.SetDataSource orsData, 3, 1

        Me.crvReporte.ReportSource = vReporte
        Me.crvReporte.ViewReport
    Else
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_RESUMEN_TOTAL]"

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adChar, adParamInput, 8, FormatoFecha(Me.dtpFecha.Value))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FILTRO", adInteger, adParamInput, , Me.cboFiltro.ListIndex)

        Dim orsData2 As ADODB.Recordset

        Set orsData2 = oCmdEjec.Execute

        Dim vReporte2   As CRAXDRT.Report

        Dim colParam2   As CRAXDRT.ParameterFieldDefinitions

        Dim objParam2   As CRAXDRT.ParameterFieldDefinition

        Dim objCrystal2 As New CRAXDRT.APPLICATION

        Set vReporte2 = objCrystal.OpenReport(PUB_RUTA_OTRO & "ReportePedidoVendedorRES.rpt")
    
        Set colParam2 = vReporte2.ParameterFields
    
        For Each objParam2 In colParam2

            Select Case objParam2.ParameterFieldName

                Case "pCant"
                    objParam2.AddCurrentValue CStr(vCAnt)

            End Select

        Next
    
        vReporte2.Database.SetDataSource orsData2, 3, 1

        Me.crvReporte.ReportSource = vReporte2
        Me.crvReporte.ViewReport

    End If

End Sub

Private Sub Form_Load()
Me.cboFiltro.ListIndex = 0
Me.dtpFecha.Value = Now
Me.chkres.Value = 0
Me.chkres.Caption = "Detallado"
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_VENDEDOR_LIST]"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

Dim orsV As ADODB.Recordset
Set orsV = oCmdEjec.Execute

Set Me.DatVendedor.RowSource = orsV
Me.DatVendedor.ListField = orsV(1).Name
Me.DatVendedor.BoundColumn = orsV(0).Name
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
    'Me.crvReporte.Top = 0
    Me.fraReporte.Left = 0
    Me.fraReporte.Height = Me.ScaleHeight - 1000
    Me.fraReporte.Width = Me.ScaleWidth
    Me.frafiltro.Height = Me.ScaleHeight - 1000
    Me.frafiltro.Width = Me.ScaleWidth
    'Me.crvReporte.Zoom 100
    
   ' Me.crvReporte.Left = 0
    Me.crvReporte.Height = Me.ScaleHeight - 1300
    Me.crvReporte.Width = Me.ScaleWidth - 100
End Sub

