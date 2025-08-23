Attribute VB_Name = "mTools"
Option Explicit
Public RES As Long
Public PUB_UNIDADS As String
Public LK_FACTORMAN As Integer
Public Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Function Devuelve_Cantidad_Marcados_LV(lv As Object)
Dim I As Integer, xdato As Integer
xdato = 0
For I = 1 To lv.ListItems.count
    If lv.ListItems.Item(I).Checked Then
    xdato = xdato + 1
    End If
Next
Devuelve_Cantidad_Marcados_LV = xdato
End Function
Public Function OpenSQLForwardOnly(ByVal strSP As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    
    ' Create the ADO objects
    Dim rs As ADODB.Recordset, cmd As ADODB.Command
    Set rs = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    ' Init the ADO objects  & the stored proc parameters
    cmd.ActiveConnection = Pub_ConnAdo
    cmd.CommandText = strSP
    cmd.CommandType = adCmdText
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly
    
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rs.ActiveConnection = Nothing
    
    Set OpenSQLForwardOnly = rs
    Exit Function

ErrHandler:
    Set OpenSQLForwardOnly = Nothing
    Set rs = Nothing
    Set cmd = Nothing
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Function

Public Sub LlenadoCbo(ByVal cbo As ComboBox, ByVal TIPREG As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = TIPREG
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cbo.ToolTipText = "TAB_TIPREG = " & TIPREG
    cbo.Clear
    Do Until tab_mayor.EOF
        cbo.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Public Sub ClearRow(ByVal pRow As Integer, ByVal grd As MSFlexGrid)
Dim pCol As Integer
Dim pCols As Integer
    pCols = grd.Cols
    For pCol = 0 To pCols - 1
        grd.TextMatrix(CLng(pRow), CLng(pCol)) = ""
    Next pCol
    grd.SetFocus
End Sub
Public Function ExistDato(ByVal Dato As Variant, ByVal aColumn As Long, ByVal aRow As Long, ByVal grd As MSFlexGrid) As Boolean
Dim iRows As Integer
Dim iRow As Long
    ExistDato = False
    iRows = grd.Rows
    For iRow = 1 To iRows - 1
        If Trim(grd.TextMatrix(iRow, aColumn)) = Dato And iRow <> aRow Then
            ExistDato = True
            Exit For
        End If
    Next iRow
End Function


Public Function Consis1(ByVal sDato As Variant) As Integer
'para cliente
    Consis1 = 0
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = val(sDato)
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
        Consis1 = 1
    End If
End Function
Public Function Consis5(ByVal sDato As Variant) As Integer
'para PROVEEDOR
    Consis5 = 0
    SQ_OPER = 1
    pu_cp = "P"
    pu_codclie = val(sDato)
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
        Consis5 = 1
    End If
End Function
Public Function consis2(ByVal sDato As Variant) As Integer
'para que no este vacio
    If Trim(sDato) = "" Then
        consis2 = 1
    End If
End Function
Public Function Consis3(ByVal grid As MSFlexGrid) As Integer
'para detalle
Dim iCount As Integer
Dim I As Integer
    For I = 1 To grid.Rows - 1
        If Trim(grid.TextMatrix(I, 13)) = "" Then
            iCount = iCount + 1
        End If
    Next I
    If ((grid.Rows - 1) = iCount) Then
        Consis3 = 1
    End If
End Function
Public Function Consis4(ByVal sDato As String) As Integer
'para fecha
'para nro de cotizacion
    If Not IsDate(sDato) Then
        Consis4 = 1
    End If
End Function
Public Sub ClearForm(Frm As Form)
On Error GoTo ErrorHandler
Dim noC%, j%
Dim c As Control
noC = Frm.count
For j = 0 To noC - 1
    Set c = Frm.Controls(j)
    If TypeOf c Is ctlText Then
       c.Text = ""
    End If
    If TypeOf c Is TextBox And c.Tag = "cls" Then
       c.Text = ""
    End If
    If TypeOf c Is ctlMaskEdBox Then
       c.Text = "__/__/____"
    End If
    If TypeOf c Is Picture Then
        'c.Picture = LoadPicture("")
    End If
    If TypeOf c Is ComboBox Then
        c.ListIndex = -1
    End If
    If TypeOf c Is label And c.Tag = "cls" Then
        c.Caption = ""
    End If
    If TypeOf c Is OptionButton Then
        c.Value = False
    End If
    If TypeOf c Is OSFindItem Then
        c.TEXTO = ""
    End If
Next
Set c = Nothing
Exit Sub
ErrorHandler:
    Set c = Nothing
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub
Public Function vString(K As Integer) As Integer
Dim sCaracter As String
Dim cE As Integer
sCaracter = UCase(Chr(K))
cE = InStr(1, "·ÈÌÛ˙Ò¡…Õ”⁄—", sCaracter)
    If ((K <= 123 And K >= 97) Or (K <= 90 And K >= 65)) Or K = 32 Or K = 8 Or cE > 0 Then
        vString = K
    Else
        vString = 0
    End If
End Function
Public Function vInteger(K As Integer) As Integer
Dim cE As Integer
Dim sCaracter As String
sCaracter = Chr(K)
cE = InStr(1, "0123456789", sCaracter)
    If cE > 0 Or K = 8 Then
        vInteger = K
    Else
        vInteger = 0
    End If
End Function
Public Function vNumeric(K As Integer) As Integer
Dim cE As Integer
Dim sCaracter As String
sCaracter = Chr(K)
cE = InStr(1, "0123456789.", sCaracter)
    If cE > 0 Or K = 8 Then
        vNumeric = K
    Else
        vNumeric = 0
    End If
End Function

Public Sub BackColorRow(ByVal iRow As Long, ByVal grd As MSFlexGrid, ByVal color As Variant)
Dim iCol As Long
    grd.Row = iRow
    For iCol = 1 To grd.Cols - 1
     grd.COL = iCol
     grd.CellBackColor = color
     'grD.CellFontBold = True
    Next
    grd.COL = 1
End Sub
Public Sub LLENADO_SUBFAM(ctlCombo As ComboBox, ByVal wfami As Integer)
On Error GoTo SALE
Dim CONTA As Integer
    CONTA = -1
    Select Case ctlCombo.Name
      Case Is = "art_subfam"
       PUB_TIPREG = 123
'      Case Is = "art_grupo"
'       PUB_TIPREG = 129
'      Case Is = "art_numero"
'       PUB_TIPREG = 130
'      Case Is = "art_linea"
'       PUB_TIPREG = 131
    End Select
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    PUB_CODART = wfami
    SQ_OPER = 3
    LEER_TAB_LLAVE
    Select Case ctlCombo.Name
      Case Is = "art_subfam"
       ctlCombo.ToolTipText = "TAB_TIPREG = 123"
      Case Is = "art_grupo"
       ctlCombo.ToolTipText = "TAB_TIPREG = 129"
      Case Is = "art_numero"
       ctlCombo.ToolTipText = "TAB_TIPREG = 130"
      Case Is = "art_linea"
       ctlCombo.ToolTipText = "TAB_TIPREG = 131"
      Case Is = "art_marca"
       ctlCombo.ToolTipText = "TAB_TIPREG = 132"
    End Select

    ctlCombo.Clear
    Do Until tab_menor.EOF
        ctlCombo.AddItem tab_menor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        DoEvents
        CONTA = CONTA + 1
        tab_menor.MoveNext
    Loop
Exit Sub
SALE:
Resume Next
End Sub

'Av Ricardo Palma cuadra 5-530

Public Sub BuscaInCbo(WCONTROL As ComboBox, txt As Integer)
Dim c As Integer
    For c = 0 To WCONTROL.ListCount - 1
        If val(Trim(Right(WCONTROL.List(c), 6))) = txt Then
            WCONTROL.ListIndex = c
            Exit Sub
        End If
    Next c
    WCONTROL.ListIndex = -1
End Sub
Public Function NumLetras() As Integer

    archi = "SELECT COUNT(CAR_NUMDOC) AS NUMLETRA FROM CARTERA WHERE CAR_CP= 'C' AND CAR_CODTRA=1455 AND CAR_CODCIA = '" & LK_CODCIA & "' "
    Set PSX = CN.CreateQuery("", archi)
    Set x = PSX.OpenResultset(rdOpenKeyset)
    x.Requery
    If x.EOF Then
        NumLetras = 0
    Else
        NumLetras = Nulo_Valor0(x("NUMLETRA"))
    End If

End Function
