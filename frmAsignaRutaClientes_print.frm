VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmAsignaRutaClientes_print 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18180
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
   ScaleHeight     =   9000
   ScaleWidth      =   18180
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer crvReporte 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18015
      DisplayGroupTree=   0   'False
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
Attribute VB_Name = "frmAsignaRutaClientes_print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oRSdatos As ADODB.Recordset
Public pVendedor As String
Public pDia As String

Private Sub Form_Load()

   
   
    Dim vReporte   As CRAXDRT.Report

    Dim colParam   As CRAXDRT.ParameterFieldDefinitions

    Dim objParam   As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal As New CRAXDRT.APPLICATION

    Set vReporte = objCrystal.OpenReport(PUB_RUTA_OTRO & "ListadoRuta.rpt")
    'Set vReporte = objCrystal.OpenReport("d:\listado.rpt")
    
    Set colParam = vReporte.ParameterFields
    
    For Each objParam In colParam
        Select Case objParam.ParameterFieldName
            Case "pVENDEDOR"
                objParam.AddCurrentValue CStr(pVendedor)
            Case "pDIA"
                objParam.AddCurrentValue CStr(pDia)
        End Select
    Next
'
    vReporte.Database.SetDataSource oRSdatos, 3, 1

    Me.crvReporte.ReportSource = vReporte
    Me.crvReporte.ViewReport

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub

'    Me.fraReporte.Left = 0
'    Me.fraReporte.Height = Me.ScaleHeight
'    Me.fraReporte.Width = Me.ScaleWidth
'    Me.frafiltro.Height = Me.ScaleHeight - 1000
'    Me.frafiltro.Width = Me.ScaleWidth
'
    Me.crvReporte.Height = Me.ScaleHeight
    Me.crvReporte.Width = Me.ScaleWidth
End Sub
