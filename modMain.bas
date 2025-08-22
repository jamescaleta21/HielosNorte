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
    
    ' Verificar que el recordset esté desconectado
    If oRS.State <> adStateClosed Then
        If Not (oRS.ActiveConnection Is Nothing) Then
            MsgBox "El Recordset no está desconectado", vbExclamation
            Exit Sub

        End If

    Else
        MsgBox "El Recordset está cerrado", vbExclamation
        Exit Sub

    End If
    
    ' Método más eficiente para eliminar todos los registros
    If oRS.RecordCount > 0 Then
        oRS.Filter = adFilterNone
        oRS.MoveFirst
        
        ' Eliminar todos los registros uno por uno
        Do Until oRS.EOF
            oRS.Delete adAffectCurrent
            oRS.MoveNext
        Loop
        
        ' Alternativa más rápida si el proveedor lo soporta:
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

    ' Asumiendo que inputDate está en formato "dd/MM/yyyy"
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

' Módulo: modValidaciones
' Propósito: Funciones de validación generales
' Función para validar fechas en formato dd/mm/yyyy
Public Function ValidarFecha(ByVal FechaTexto As String, Optional ByVal PermitirVacio As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar si está vacío y si se permite
    If Trim(FechaTexto) = "" Or FechaTexto = "__/__/____" Then
        ValidarFecha = PermitirVacio
        Exit Function
    End If
    
    ' Verificar longitud mínima (dd/mm/yyyy = 10 caracteres)
    If Len(Trim(FechaTexto)) < 10 Then
        ValidarFecha = False
        Exit Function
    End If
    
    ' Reemplazar caracteres de máscara si existen
    FechaTexto = Replace(FechaTexto, "_", "0")
    
    ' Extraer día, mes y año
    Dim DIA As Integer, MES As Integer, año As Integer
    DIA = CInt(Left(FechaTexto, 2))
    MES = CInt(Mid(FechaTexto, 4, 2))
    año = CInt(Right(FechaTexto, 4))
    
    ' Validación básica de rangos
    If año < 100 Or año > 9999 Then Exit Function  ' Año entre 0100 y 9999
    If MES < 1 Or MES > 12 Then Exit Function     ' Mes entre 1 y 12
    If DIA < 1 Or DIA > 31 Then Exit Function     ' Día entre 1 y 31
    
    ' Validación de meses con 30 días
    If (MES = 4 Or MES = 6 Or MES = 9 Or MES = 11) And DIA > 30 Then Exit Function
    
    ' Validación especial para febrero
    If MES = 2 Then
        ' Verificar año bisiesto
        If (año Mod 4 = 0 And año Mod 100 <> 0) Or (año Mod 400 = 0) Then
            If DIA > 29 Then Exit Function
        Else
            If DIA > 28 Then Exit Function
        End If
    End If
    
    ' Intentar convertir a fecha para validación final
    Dim FechaCompleta As Date
    FechaCompleta = DateSerial(año, MES, DIA)
    
    ' Si llegó hasta aquí, la fecha es válida
    ValidarFecha = True
    Exit Function
    
ErrorHandler:
    ValidarFecha = False
End Function
