VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResumenDiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resúmen Diario"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResumenDiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   12105
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   585
         Left            =   10320
         Picture         =   "frmResumenDiario.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   170
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   320
         Left            =   3480
         TabIndex        =   16
         Top             =   302
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102367233
         CurrentDate     =   42703
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Documento:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1320
         TabIndex        =   17
         Top             =   365
         Width           =   2100
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   11895
      Begin VB.CheckBox chkMarca 
         Caption         =   "Marcar Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvData 
         Height          =   3495
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   11895
      Begin VB.TextBox txtSec 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   7080
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "1"
         Top             =   177
         Width           =   615
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
         Height          =   600
         Left            =   10320
         Picture         =   "frmResumenDiario.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   577
         Width           =   1215
      End
      Begin VB.CommandButton cmdCarpeta 
         Height          =   360
         Left            =   9360
         Picture         =   "frmResumenDiario.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Cambie Carpeta"
         Top             =   817
         Width           =   375
      End
      Begin VB.CommandButton cmdRutaDefecto 
         Height          =   360
         Left            =   9840
         Picture         =   "frmResumenDiario.frx":0E28
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Carpeta por Defecto"
         Top             =   817
         Width           =   375
      End
      Begin MSComCtl2.UpDown udSec 
         Height          =   315
         Left            =   7680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   177
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSec"
         BuddyDispid     =   196615
         OrigLeft        =   5280
         OrigTop         =   480
         OrigRight       =   5520
         OrigBottom      =   975
         Max             =   999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpFechaReporte 
         Height          =   320
         Left            =   2040
         TabIndex        =   0
         Top             =   177
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102367233
         CurrentDate     =   42703
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Reporte:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Secuencia:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5160
         TabIndex        =   12
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guardar en:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lblRuta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   9195
      End
   End
End
Attribute VB_Name = "frmResumenDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oRSdata As ADODB.Recordset
Private itemM As Integer

Private Sub CrearArchivoPlano()

Dim CI          As Integer

Dim sCadena     As String

Dim cont        As Integer

Dim sEPARADOR   As String

Dim obj_FSO     As Object

Dim Archivo     As Object
Dim ArchivoTRD As Object

Dim sARCHIVOrdi As String
Dim sARCHIVOtrd As String

Dim sRUC        As String

Dim sCCC        As String

Dim cItem       As Integer
    

If itemM = 0 Then
    MsgBox "No ha marcado ningun documento.", vbCritical, Pub_Titulo
    Exit Sub
End If
    
On Error GoTo Procesar



sEPARADOR = "|"

sCadena = ""

cont = 1
    
sCCC = Right("000" & Me.txtSec.Text, 3)
    
If LK_CODCIA = "01" Then
sRUC = Leer_Ini(App.Path & "\config.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "02" Then
sRUC = Leer_Ini(App.Path & "\config2.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "03" Then
sRUC = Leer_Ini(App.Path & "\config3.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "04" Then
sRUC = Leer_Ini(App.Path & "\config4.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "05" Then
sRUC = Leer_Ini(App.Path & "\config5.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "06" Then
sRUC = Leer_Ini(App.Path & "\config6.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "07" Then
sRUC = Leer_Ini(App.Path & "\config7.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "08" Then
sRUC = Leer_Ini(App.Path & "\config8.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "09" Then
sRUC = Leer_Ini(App.Path & "\config9.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "10" Then
sRUC = Leer_Ini(App.Path & "\config10.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "11" Then
sRUC = Leer_Ini(App.Path & "\config11.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "12" Then
sRUC = Leer_Ini(App.Path & "\config12.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "13" Then
sRUC = Leer_Ini(App.Path & "\config13.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "14" Then
sRUC = Leer_Ini(App.Path & "\config14.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "15" Then
sRUC = Leer_Ini(App.Path & "\config15.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "16" Then
sRUC = Leer_Ini(App.Path & "\config16.ini", "RUC", "C:\")
ElseIf LK_CODCIA = "17" Then
sRUC = Leer_Ini(App.Path & "\config17.ini", "RUC", "C:\")
Else
End If
    
'sARCHIVOrdi = sRUC & "-RC-" & CStr(Year(Me.dtpFechaReporte.Value)) + Right("00" & CStr(Month(Me.dtpFechaReporte.Value)), 2) + Right("00" & CStr(Day(Me.dtpFechaReporte.Value)), 2) + "-" & sCCC & ".rdi"
sARCHIVOrdi = sRUC & "-RC-" & CStr(Year(Me.dtpFechaReporte.Value)) + Right("00" & CStr(Month(Me.dtpFechaReporte.Value)), 2) + Right("00" & CStr(Day(Me.dtpFechaReporte.Value)), 2) + "-" & sCCC & ".rdi"
sARCHIVOtrd = sRUC & "-RC-" & CStr(Year(Me.dtpFechaReporte.Value)) + Right("00" & CStr(Month(Me.dtpFechaReporte.Value)), 2) + Right("00" & CStr(Day(Me.dtpFechaReporte.Value)), 2) + "-" & sCCC & ".trd"
    
Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
'Set Archivo = obj_FSO.CreateTextFile(Me.LBLRUTA.Caption + sARCHIVOrdi, True)
Set Archivo = obj_FSO.CreateTextFile(Me.lblRuta.Caption + sARCHIVOrdi, True)
Set ArchivoTRD = obj_FSO.CreateTextFile(Me.lblRuta.Caption + sARCHIVOtrd, True)
'Set Archivo = obj_FSO.CreateTextFile("D:\APP_FE\Service_WCF_8500\FileServer\Uploads\" + sARCHIVO, True)

CI = 1

If oRSdata.EOF Then oRSdata.MoveFirst

For cItem = 1 To Me.lvData.ListItems.count

    If Me.lvData.ListItems.Item(cItem).Checked Then
        oRSdata.Filter = "IDOCTO='" & Me.lvData.ListItems.Item(cItem).SubItems(2) & "' " & "AND INDICE= " & Me.lvData.ListItems.Item(cItem).SubItems(7)

        If Not oRSdata.EOF Then
            sCadena = sCadena & CStr(oRSdata!fechadocto) + sEPARADOR + CStr(oRSdata!FECHACTUAL) + sEPARADOR + oRSdata!tipodocto + sEPARADOR + oRSdata!IDOCTO + sEPARADOR + oRSdata!TDI + sEPARADOR & oRSdata!NRODOCUSUARIO & sEPARADOR & oRSdata!moneda & sEPARADOR & oRSdata!CAMPO1 & sEPARADOR & oRSdata!Total & sEPARADOR & FormatNumber(oRSdata!EXO, 2) & sEPARADOR & FormatNumber(oRSdata!INA, 2) & sEPARADOR & FormatNumber(oRSdata!GRA, 2) & sEPARADOR & FormatNumber(oRSdata!icbper, 2) & sEPARADOR & oRSdata!TOTALVTA & sEPARADOR & oRSdata!TIPDOCTOMODIFICA & sEPARADOR & oRSdata!SERIEBOLMODIFICA & sEPARADOR & oRSdata!NROBOLMODIFICA & sEPARADOR & oRSdata!REGPERCEPCION & sEPARADOR & oRSdata!PORCPERCEPCION & sEPARADOR & oRSdata!BASEIMPERCEPCION & sEPARADOR & oRSdata!MONTOPERCEPCION & sEPARADOR & oRSdata!MONTOTOTINCPERCEPCION & sEPARADOR & oRSdata!ESTADO & sEPARADOR
            
            If CI < itemM Then
                sCadena = sCadena & vbCrLf

            End If
            
        End If

        CI = CI + 1
    End If

Next


'Escribimos lineas
Archivo.WriteLine sCadena
    
'Cerramos el fichero
Archivo.Close
Set Archivo = Nothing
oRSdata.Filter = ""
oRSdata.MoveFirst
Dim orsTRD As New ADODB.Recordset
Set orsTRD = oRSdata.NextRecordset
orsTRD.Filter = ""

sCadena = ""
Dim i As Integer
CI = 1
Do While Not orsTRD.EOF
    sCadena = sCadena & orsTRD!c0 & sEPARADOR & orsTRD!c1 & sEPARADOR & orsTRD!c2 & sEPARADOR & orsTRD!c3 & sEPARADOR & orsTRD!c4 & sEPARADOR & orsTRD!c5 & sEPARADOR
    If CI < orsTRD.RecordCount Then sCadena = sCadena & vbCrLf
    CI = CI + 1
    orsTRD.MoveNext
Loop

ArchivoTRD.WriteLine sCadena
ArchivoTRD.Close
Set ArchivoTRD = Nothing

Set orsTRD = Nothing
oRSdata.Requery
MsgBox "Archivo creado correctamente", vbInformation, Pub_Titulo

Exit Sub

Procesar:
MsgBox "Error al crear el archivo", vbCritical, Pub_Titulo
End Sub

Private Sub chkMarca_Click()
Dim i As Integer
For i = 1 To Me.lvData.ListItems.count
    Me.lvData.ListItems.Item(i).Checked = Me.chkMarca.Value
Next
End Sub

Private Sub cmdCarpeta_Click()
 Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
'    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
'  If Len(Trim(ret)) = 0 Then Exit Sub
'    Me.lblRuta.Caption = ret & "\"
End Sub

Private Sub cmdGenerar_Click()
If Len(Trim(Me.lblRuta.Caption)) = 0 Then
    MsgBox "Debe especificar la ruta del archivo a generar.", vbCritical, Pub_Titulo
    Exit Sub
End If
If Not IsNumeric(Me.txtSec.Text) Then
    MsgBox "El Nro de Secuencia es incorrecto.", vbCritical, Pub_Titulo
    Me.txtSec.SetFocus
    Exit Sub
End If
If Val(Me.txtSec.Text) <= 0 Then
    MsgBox "El Nro de Secuencia es incorrecto.", vbCritical, Pub_Titulo
    Me.txtSec.SetFocus
    Exit Sub
End If
If Me.lvData.ListItems.count = 0 Then
    MsgBox "No hay ningun documento para generar la información.", vbInformation, Pub_Titulo
    Exit Sub
End If
itemM = Devuelve_Cantidad_Marcados_LV(Me.lvData)
If itemM = 0 Then
    MsgBox "Debe marca al menos un documento", vbInformation, Pub_Titulo
    Exit Sub
End If

CrearArchivoPlano
End Sub

Private Sub cmdmostrar_Click()

Me.lvData.ListItems.Clear

LimpiaParametros oCmdEjec
    
oCmdEjec.CommandText = "SP_RESUMEN_DIARIO"
    
oCmdEjec.CommandType = adCmdStoredProc
'  Exit Sub
    
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAACTUAL", adDBTimeStamp, adParamInput, , Me.dtpFechaReporte)
    
Set oRSdata = oCmdEjec.Execute
    
If oRSdata.RecordCount <> 0 Then

    Do While Not oRSdata.EOF
        Set itemX = Me.lvData.ListItems.Add(, , Trim(oRSdata!fechadocto))
        itemX.SubItems(1) = Trim(oRSdata!tipodocto)
        itemX.SubItems(2) = oRSdata!IDOCTO
        itemX.SubItems(3) = oRSdata!cliente
        itemX.SubItems(4) = oRSdata!CAMPO1
        itemX.SubItems(5) = oRSdata!icbper
        itemX.SubItems(6) = oRSdata!TOTALVTA
        itemX.SubItems(7) = oRSdata!indice 'LINEA NUEVA 04-10-18
        itemX.SubItems(8) = oRSdata!ESTADO 'LINEA NUEVA 06-10-18
        oRSdata.MoveNext
    Loop


    

Else
    MsgBox "No se ha encontrado ningun documento en la fecha proporcionada.", vbInformation, Pub_Titulo
End If

End Sub

Private Sub cmdRutaDefecto_Click()
If LK_CODCIA = "01" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "02" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "03" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "04" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "05" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config5.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "06" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config6.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "07" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config7.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "08" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config8.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "09" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config9.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "10" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config10.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "11" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config11.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "12" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config12.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "13" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config13.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "14" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config14.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "15" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config15.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "16" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config16.ini", "CARPETA", "C:\")
ElseIf LK_CODCIA = "17" Then
Me.lblRuta.Caption = Leer_Ini(App.Path & "\config17.ini", "CARPETA", "C:\")
Else
End If
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
dtpFecha.Value = LK_FECHA_DIA
Me.dtpFechaReporte.Value = LK_FECHA_DIA
cmdRutaDefecto_Click
ConfigurarLV
End Sub

Private Sub ConfigurarLV()
With Me.lvData
    .ColumnHeaders.Add , , "FECHA", 1500
    .ColumnHeaders.Add , , "TIPO DOCTO"
    .ColumnHeaders.Add , , "NRO DOCTO"
    .ColumnHeaders.Add , , "CLIENTE", 3000
    .ColumnHeaders.Add , , "SUB TOTAL"
    .ColumnHeaders.Add , , "IGV"
    .ColumnHeaders.Add , , "TOTAL", 1000
    .ColumnHeaders.Add , , "INDICE", 0
    .ColumnHeaders.Add , , "ESTADO"
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = False
    .View = lvwReport
    .HideSelection = False
End With
End Sub


