Attribute VB_Name = "modMain"
Public Const cClave As String = "anteromariano"
Public oRSmain As ADODB.Recordset

Public Function devuelveIDempresaXdefecto() As Integer
    On Error GoTo defecto

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_EMPRESA_PORDEFECTO]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPRESA", adInteger, adParamOutput, , 0)
    oCmdEjec.Execute
    devuelveIDempresaXdefecto = oCmdEjec.Parameters("@IDEMPRESA").Value
    CerrarConexion False
    Exit Function
defecto:
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Function

Public Sub LimpiarMaskEdBox(oMask As String, objControl As MaskEdBox)
objControl.Mask = ""
objControl.Text = ""
objControl.Mask = oMask
End Sub


Public Sub EliminarRegistrosRecordSet(oRS As ADODB.Recordset)

    On Error GoTo ErrorHandler
    
    ' Verificar que el recordset est� desconectado
    If oRS.State <> adStateClosed Then
        If Not (oRS.ActiveConnection Is Nothing) Then
            MsgBox "El Recordset no est� desconectado", vbExclamation
            Exit Sub

        End If

    Else
        MsgBox "El Recordset est� cerrado", vbExclamation
        Exit Sub

    End If
    
    ' M�todo m�s eficiente para eliminar todos los registros
    If oRS.RecordCount > 0 Then
        oRS.Filter = adFilterNone
        oRS.MoveFirst
        
        ' Eliminar todos los registros uno por uno
        Do Until oRS.EOF
            oRS.Delete adAffectCurrent
            oRS.MoveNext
        Loop
        
        ' Alternativa m�s r�pida si el proveedor lo soporta:
        ' rs.Requery
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al eliminar registros: " & Err.Description, vbExclamation

End Sub

Function ConvertirFechaFormat_yyyyMMdd(ByVal inputDate As String) As String

    Dim dayPart    As String

    Dim monthPart  As String

    Dim yearPart   As String

    Dim outputDate As String

    ' Asumiendo que inputDate est� en formato "dd/MM/yyyy"
    dayPart = Mid(inputDate, 1, 2)
    monthPart = Mid(inputDate, 4, 2)
    yearPart = Mid(inputDate, 7, 4)

    ' Construir la nueva fecha en formato "yyyyMMdd"
    outputDate = yearPart & monthPart & dayPart

    ' Devolver la nueva fecha
    ConvertirFechaFormat_yyyyMMdd = outputDate

End Function

Public Sub HandleEnterKey(KeyCode As Integer, NextControl As Control)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        NextControl.SetFocus
        
        ' Intenta seleccionar el texto en cualquier control que lo soporte
        On Error Resume Next
        
        Dim ctl As Object
        Set ctl = NextControl
        
        ' Verifica si el control tiene las propiedades SelStart y SelLength
        If Not ctl.SelStart Is Nothing And Not ctl.SelLength Is Nothing Then
            ctl.SelStart = 0
            ctl.SelLength = Len(ctl.Text)
        End If
        
        On Error GoTo 0
    End If
End Sub

' M�dulo: modValidaciones
' Prop�sito: Funciones de validaci�n generales
' Funci�n para validar fechas en formato dd/mm/yyyy
Public Function ValidarFecha(ByVal FechaTexto As String, Optional ByVal PermitirVacio As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar si est� vac�o y si se permite
    If Trim(FechaTexto) = "" Or FechaTexto = "__/__/____" Then
        ValidarFecha = PermitirVacio
        Exit Function
    End If
    
    ' Verificar longitud m�nima (dd/mm/yyyy = 10 caracteres)
    If Len(Trim(FechaTexto)) < 10 Then
        ValidarFecha = False
        Exit Function
    End If
    
    ' Reemplazar caracteres de m�scara si existen
    FechaTexto = Replace(FechaTexto, "_", "0")
    
    ' Extraer d�a, mes y a�o
    Dim DIA As Integer, MES As Integer, a�o As Integer
    DIA = CInt(Left(FechaTexto, 2))
    MES = CInt(Mid(FechaTexto, 4, 2))
    a�o = CInt(Right(FechaTexto, 4))
    
    ' Validaci�n b�sica de rangos
    If a�o < 100 Or a�o > 9999 Then Exit Function  ' A�o entre 0100 y 9999
    If MES < 1 Or MES > 12 Then Exit Function     ' Mes entre 1 y 12
    If DIA < 1 Or DIA > 31 Then Exit Function     ' D�a entre 1 y 31
    
    ' Validaci�n de meses con 30 d�as
    If (MES = 4 Or MES = 6 Or MES = 9 Or MES = 11) And DIA > 30 Then Exit Function
    
    ' Validaci�n especial para febrero
    If MES = 2 Then
        ' Verificar a�o bisiesto
        If (a�o Mod 4 = 0 And a�o Mod 100 <> 0) Or (a�o Mod 400 = 0) Then
            If DIA > 29 Then Exit Function
        Else
            If DIA > 28 Then Exit Function
        End If
    End If
    
    ' Intentar convertir a fecha para validaci�n final
    Dim FechaCompleta As Date
    FechaCompleta = DateSerial(a�o, MES, DIA)
    
    ' Si lleg� hasta aqu�, la fecha es v�lida
    ValidarFecha = True
    Exit Function
    
ErrorHandler:
    ValidarFecha = False
End Function
