VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFormGenPedidosVD_Consulta 
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
   Icon            =   "frmFormGenPedidosVD_Consulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   16560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraReporte 
      Height          =   7935
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   16335
      Begin CRVIEWERLibCtl.CRViewer crvReporte 
         Height          =   7575
         Left            =   120
         TabIndex        =   4
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
      TabIndex        =   5
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
         Left            =   8520
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   270
         Width           =   1935
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar"
         Height          =   600
         Left            =   14760
         Picture         =   "frmFormGenPedidosVD_Consulta.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   6120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   79757313
         CurrentDate     =   44900
      End
      Begin MSDataListLib.DataCombo DatVendedor 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   5400
         TabIndex        =   8
         Top             =   300
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   300
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmFormGenPedidosVD_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pIDempresa As Integer
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
    
    Dim orsTEMP1   As New ADODB.Recordset

    Dim vReporte   As CRAXDRT.Report

    Dim colParam   As CRAXDRT.ParameterFieldDefinitions

    Dim objParam   As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal As New CRAXDRT.APPLICATION
    
    MousePointer = vbHourglass

    On Error GoTo cMostrar

    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_PEDIDOS_VENDEDOR]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adChar, adParamInput, 8, FormatoFecha(Me.dtpFecha.Value))

    Dim orsc As ADODB.Recordset

    Set orsc = oCmdEjec.Execute

    Dim vCAnt As Integer

    vCAnt = orsc!cant

    If chkres.Value = 0 Then
        LimpiaParametros oCmdEjec, True
        oCmdEjec.CommandText = "[vd].[USP_PEDIDO_RESUMEN]"

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adChar, adParamInput, 8, FormatoFecha(Me.dtpFecha.Value))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, 2, pIDempresa)

        Set oRSmain = oCmdEjec.Execute
    
        Set vReporte = objCrystal.OpenReport(PUB_RUTA_OTRO & "ReportePedidoVendedorVD.rpt")
    
        Set colParam = vReporte.ParameterFields
    
        For Each objParam In colParam

            Select Case objParam.ParameterFieldName

                Case "pCant"
                    objParam.AddCurrentValue CStr(vCAnt)

            End Select

        Next

If oRSmain.RecordCount <> 0 Then
        orsTEMP1.CursorLocation = adUseClient
        orsTEMP1.Fields.Append oRSmain(0).Name, oRSmain(0).Type, oRSmain(0).DefinedSize
        orsTEMP1.Fields.Append oRSmain(1).Name, oRSmain(1).Type, oRSmain(1).DefinedSize
        orsTEMP1.Fields.Append oRSmain(2).Name, oRSmain(2).Type, oRSmain(2).DefinedSize
        orsTEMP1.Fields.Append oRSmain(3).Name, oRSmain(3).Type, oRSmain(3).DefinedSize
        orsTEMP1.Fields.Append oRSmain(4).Name, oRSmain(4).Type, oRSmain(4).DefinedSize
        orsTEMP1.Fields.Append oRSmain(5).Name, oRSmain(5).Type, oRSmain(5).DefinedSize
        orsTEMP1.Fields.Append oRSmain(6).Name, oRSmain(6).Type, oRSmain(6).DefinedSize
        orsTEMP1.Open
        
        
    
        ' Copiar datos del Recordset
        oRSmain.MoveFirst

        Do Until oRSmain.EOF
            orsTEMP1.AddNew
            orsTEMP1.Fields(0).Value = oRSmain.Fields(0).Value
            orsTEMP1.Fields(1).Value = oRSmain.Fields(1).Value
            orsTEMP1.Fields(2).Value = oRSmain.Fields(2).Value
            orsTEMP1.Fields(3).Value = oRSmain.Fields(3).Value
            orsTEMP1.Fields(4).Value = oRSmain.Fields(4).Value
            orsTEMP1.Fields(5).Value = oRSmain.Fields(5).Value
            orsTEMP1.Fields(6).Value = oRSmain.Fields(6).Value
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
    Else
        LimpiaParametros oCmdEjec, True
        oCmdEjec.CommandText = "[vd].[USP_PEDIDO_RESUMEN_TOTAL]"

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDVENDEDOR", adInteger, adParamInput, , Me.DatVendedor.BoundText)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adChar, adParamInput, 8, FormatoFecha(Me.dtpFecha.Value))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamInput, , pIDempresa)

        Set oRSmain = oCmdEjec.Execute

        Set vReporte = objCrystal.OpenReport(PUB_RUTA_OTRO & "ReportePedidoVendedorVDRES.rpt")
    
        Set colParam = vReporte.ParameterFields
    
        For Each objParam In colParam

            Select Case objParam.ParameterFieldName

                Case "pUSER"
                    objParam.AddCurrentValue CStr(LK_CODUSU)

            End Select

        Next
        
        If oRSmain.RecordCount <> 0 Then
        
        orsTEMP1.CursorLocation = adUseClient
        orsTEMP1.Fields.Append oRSmain(0).Name, oRSmain(0).Type, oRSmain(0).DefinedSize
        orsTEMP1.Fields.Append oRSmain(1).Name, oRSmain(1).Type, oRSmain(1).DefinedSize
        orsTEMP1.Fields.Append oRSmain(2).Name, oRSmain(2).Type, oRSmain(2).DefinedSize
        orsTEMP1.Fields.Append oRSmain(3).Name, oRSmain(3).Type, oRSmain(3).DefinedSize
        orsTEMP1.Fields.Append oRSmain(4).Name, oRSmain(4).Type, oRSmain(4).DefinedSize
        orsTEMP1.Open
    
        ' Copiar datos del Recordset
        oRSmain.MoveFirst

        Do Until oRSmain.EOF
            orsTEMP1.AddNew
            orsTEMP1.Fields(0).Value = oRSmain.Fields(0).Value
            orsTEMP1.Fields(1).Value = oRSmain.Fields(1).Value
            orsTEMP1.Fields(2).Value = oRSmain.Fields(2).Value
            orsTEMP1.Fields(3).Value = oRSmain.Fields(3).Value
            orsTEMP1.Fields(4).Value = oRSmain.Fields(4).Value
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

    End If

    MousePointer = vbDefault
    CerrarConexion True
    Exit Sub
cMostrar:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

    On Error GoTo cCarga

    MousePointer = vbHourglass
    pIDempresa = devuelveIDempresaXdefecto
    
    Me.dtpFecha.Value = Now
    Me.chkres.Value = 0
    Me.chkres.Caption = "Detallado"
    LimpiaParametros oCmdEjec, True
    oCmdEjec.CommandText = "[vd].[USP_VENDEDOR_LIST]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
    Dim orsTEMP1 As New ADODB.Recordset

    Set oRSmain = oCmdEjec.Execute
    
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
    
    ' Configurar DataVendedor
    Set Me.DatVendedor.RowSource = orsTEMP1
    Me.DatVendedor.ListField = orsTEMP1.Fields(1).Name
    Me.DatVendedor.BoundColumn = orsTEMP1.Fields(0).Name
    Me.DatVendedor.BoundText = -1
    CerrarConexion True
    MousePointer = vbDefault
    Exit Sub
cCarga:
    MousePointer = vbDefault
    CerrarConexion True
    MsgBox Err.Description, vbCritical, Pub_Titulo

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

