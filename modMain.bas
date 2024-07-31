Attribute VB_Name = "modMain"
Public Const cClave As String = "accesodenegado$1"

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
