VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMetaReporte 
   Caption         =   "Listado de Metas"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMetaReporte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   16215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraReporte 
      Height          =   7095
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   15975
      Begin CRVIEWERLibCtl.CRViewer crvReporte 
         Height          =   6735
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   15735
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
   Begin VB.Frame FraFiltro 
      Height          =   1095
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   15975
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   720
         Left            =   14520
         Picture         =   "frmMetaReporte.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   13080
         Picture         =   "frmMetaReporte.frx":1054
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optfiltro 
         Caption         =   "Todos"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optfiltro 
         Caption         =   "Inactivo"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optfiltro 
         Caption         =   "Activo"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMetaReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pFiltro As Integer
Private pIDempresa As Integer

Private Sub cmdBuscar_Click()
cargarReporte
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
pIDempresa = devuelveIDempresaXdefecto
cargarReporte
End Sub

Private Sub cargarReporte()

    
    Dim orsTEMP1   As New ADODB.Recordset

    Dim vReporte   As CRAXDRT.Report

    Dim colParam   As CRAXDRT.ParameterFieldDefinitions

    Dim objParam   As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal As New CRAXDRT.APPLICATION
    
    MousePointer = vbHourglass

    On Error GoTo cMostrar

    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[dbo].[USP_META_LIST]"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adInteger, adParamInput, , Me.optfiltro(pFiltro).Index)
    

    Set oRSmain = oCmdEjec.Execute
    
    Set vReporte = objCrystal.OpenReport(PUB_RUTA_OTRO & "ListadoMeta.rpt")
    
    Set colParam = vReporte.ParameterFields
    
    For Each objParam In colParam

        Select Case objParam.ParameterFieldName

            Case "pUSER"
                objParam.AddCurrentValue CStr(LK_CODUSU)

'            Case "pVENDEDOR"
'                objParam.AddCurrentValue CStr(Me.DatVendedor.Text)

        End Select

    Next

    If oRSmain.RecordCount <> 0 Then
        orsTEMP1.CursorLocation = adUseClient
        orsTEMP1.Fields.Append oRSmain(0).Name, oRSmain(0).Type, oRSmain(0).DefinedSize
        orsTEMP1.Fields.Append oRSmain(1).Name, oRSmain(1).Type, oRSmain(1).DefinedSize
        orsTEMP1.Fields.Append oRSmain(2).Name, oRSmain(2).Type, oRSmain(2).DefinedSize
        orsTEMP1.Fields.Append oRSmain(3).Name, oRSmain(3).Type, oRSmain(3).DefinedSize
        orsTEMP1.Open
    
        ' Copiar datos del Recordset
        oRSmain.MoveFirst

        Do Until oRSmain.EOF
            orsTEMP1.AddNew
            orsTEMP1.Fields(0).Value = oRSmain.Fields(0).Value
            orsTEMP1.Fields(1).Value = oRSmain.Fields(1).Value
            orsTEMP1.Fields(2).Value = oRSmain.Fields(2).Value
            orsTEMP1.Fields(3).Value = oRSmain.Fields(3).Value
            orsTEMP1.Update
            oRSmain.MoveNext
        Loop
        
        vReporte.Database.SetDataSource orsTEMP1, 3, 1

        Me.crvReporte.ReportSource = vReporte
        Me.crvReporte.ViewReport
    Else
        vReporte.Database.SetDataSource oRSmain, 3, 1
        Me.crvReporte.ReportSource = vReporte
        Me.crvReporte.ViewReport
        MsgBox "No se encontraron registros.", vbInformation, Pub_Titulo

    End If

    MousePointer = vbDefault
    CerrarConexion True
    Exit Sub
cMostrar:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub
    'Me.crvReporte.Top = 0
    Me.FraReporte.Left = Me.FraFiltro.Left
    Me.FraReporte.Height = Me.ScaleHeight - 1000
    Me.FraReporte.Width = Me.ScaleWidth
'    Me.frafiltro.Height = Me.ScaleHeight - 1000
   Me.FraFiltro.Width = Me.FraReporte.Width - 100
    'Me.crvReporte.Zoom 100
     Me.cmdCerrar.Left = (Me.FraReporte.Width - Me.cmdCerrar.Width) - 200
     Me.cmdBuscar.Left = (Me.FraReporte.Width - Me.cmdCerrar.Width) - Me.cmdCerrar.Width - 300
    ' Me.crvReporte.Left = 0
    Me.crvReporte.Height = Me.ScaleHeight - 1300
    Me.crvReporte.Width = Me.ScaleWidth - 200
End Sub

Private Sub optfiltro_Click(Index As Integer)
pFiltro = Index
End Sub
