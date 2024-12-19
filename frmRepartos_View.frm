VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmRepartos_View 
   Caption         =   "Form1"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20895
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
   ScaleHeight     =   10110
   ScaleWidth      =   20895
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTReportes 
      Height          =   9855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20655
      _ExtentX        =   36433
      _ExtentY        =   17383
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Reporte Detallado"
      TabPicture(0)   =   "frmRepartos_View.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "crVisorDet"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Reporte Resumen"
      TabPicture(1)   =   "frmRepartos_View.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "crVisorRes"
      Tab(1).ControlCount=   1
      Begin CRVIEWERLibCtl.CRViewer crVisorDet 
         Height          =   9255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   20415
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
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin CRVIEWERLibCtl.CRViewer crVisorRes 
         Height          =   9255
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   20415
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
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
   End
End
Attribute VB_Name = "frmRepartos_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pRSdata As ADODB.Recordset
Public pRSdataRes As ADODB.Recordset
Public pCantidad As Integer
Public pIDREPARTO As Integer
Public pREPARTIDOR As String
Public pOBS As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

    Dim vReporte   As CRAXDRT.Report

    Dim colParam   As CRAXDRT.ParameterFieldDefinitions

    Dim objParam   As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal As New CRAXDRT.APPLICATION

    Set vReporte = objCrystal.OpenReport(PUB_RUTA_OTRO & "ReporteReparto.rpt")
    
    Set colParam = vReporte.ParameterFields
    
    For Each objParam In colParam

        Select Case objParam.ParameterFieldName

            Case "pCant"
                objParam.AddCurrentValue CStr(pCantidad)

            Case "pREPARTO"
                objParam.AddCurrentValue "REPARTO NRO " & CStr(pIDREPARTO)

            Case "pREPARTIDOR"
                objParam.AddCurrentValue pREPARTIDOR

            Case "pOBS"
                objParam.AddCurrentValue pOBS

        End Select

    Next
    
    vReporte.Database.SetDataSource pRSdata, 3, 1

    Me.crVisorDet.ReportSource = vReporte
    Me.crVisorDet.ViewReport
    
    '--------------
    Set vReporte = objCrystal.OpenReport(PUB_RUTA_OTRO & "ReportePedidoRepartidorRES.rpt")
    
    Set colParam = vReporte.ParameterFields
    
    For Each objParam In colParam

        Select Case objParam.ParameterFieldName

            Case "pCant"
                objParam.AddCurrentValue CStr(pCantidad)

            Case "pREPARTO"
                objParam.AddCurrentValue "REPARTO NRO " & CStr(pIDREPARTO)

            Case "pOBS"
                objParam.AddCurrentValue pOBS

        End Select

    Next
    
    vReporte.Database.SetDataSource pRSdataRes, 3, 1

    Me.crVisorRes.ReportSource = vReporte
    Me.crVisorRes.ViewReport

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    
    Me.SSTReportes.Height = Me.ScaleHeight - 150
    Me.SSTReportes.Width = Me.ScaleWidth - 150
 
    Me.crVisorDet.Height = Me.ScaleHeight - 800
    Me.crVisorDet.Width = Me.ScaleWidth - 500
    
    Me.crVisorRes.Height = Me.ScaleHeight - 800
    Me.crVisorRes.Width = Me.ScaleWidth - 500

End Sub

