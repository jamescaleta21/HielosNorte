VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPedidoHorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horario Permitido para pasar Pedidos"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5160
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
   ScaleHeight     =   1845
   ScaleWidth      =   5160
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   600
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm:ss"" h"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   202375170
      CurrentDate     =   45148
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   202375170
      CurrentDate     =   45148
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   600
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmPedidoHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBorrar_Click()
On Error GoTo cElimina
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_HORARIO_DELETE]"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)

Dim orsData As ADODB.Recordset
Set orsData = oCmdEjec.Execute

Dim mensaje() As String
If Not orsData.EOF Then
    mensaje = Split(orsData.Fields(0).Value, "=")
    If mensaje(0) = 0 Then
        MsgBox mensaje(1), vbInformation, Pub_Titulo
        Me.LblMensaje.Caption = "No se ha configurado Horario"
        Me.dtpDesde.Value = "00:00:00"
        Me.dtpHasta.Value = "00:00:00"
    Else
        MsgBox mensaje(1), vbCritical, Pub_Titulo
    End If
End If


Exit Sub
cElimina:
MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub cmdGrabar_Click()

If Me.dtpDesde.Value > Me.dtpHasta.Value Then
    MsgBox "Horas incorrectas.", vbCritical, Pub_Titulo
    Exit Sub
End If

    On Error GoTo cSave

Dim dd()  As String
dd = Split(CStr(Me.dtpDesde.Value), " ")

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_HORARIO_REGISTER]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@INI", adVarChar, adParamInput, 10, Format(Me.dtpDesde.Value, "HH:mm:ss"))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fin", adVarChar, adParamInput, 10, Format(Me.dtpHasta.Value, "HH:mm:ss"))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)


    Dim orsData As ADODB.Recordset

    Set orsData = oCmdEjec.Execute

    Dim mensaje() As String

    If Not orsData.EOF Then
        mensaje = Split(orsData.Fields(0), "=")

        If mensaje(0) = 0 Then
        Me.LblMensaje.Caption = "Horario actual configurado:"
            MsgBox mensaje(1), vbInformation, Pub_Titulo
        Else
            MsgBox mensaje(1), vbCritical, Pub_Titulo

        End If

    End If

    Exit Sub
cSave:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
MostrarDatos
'Me.dtpDesde.Format = dtpCustom
'Me.dtpDesde.CustomFormat = "HH:mm:ss"
End Sub

Private Sub MostrarDatos()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_PEDIDO_HORARIO_LOAD]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    Dim orsData As ADODB.Recordset

    Set orsData = oCmdEjec.Execute

    If Not orsData.EOF Then
        Me.LblMensaje.Caption = "Horario actual configurado:"
        Me.dtpDesde.Value = orsData!ini
        Me.dtpHasta.Value = orsData!fin
    Else
        Me.LblMensaje.Caption = "No se ha configurado Horario"

    End If

End Sub
