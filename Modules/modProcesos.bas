Attribute VB_Name = "modProcesos"
Public gNumeroPedido As String 'Número de Pedido a Anular
Public gPosicionPedido As Integer 'Determina la posición del pedido a anular

Public Function AnularPedido() As Boolean

    On Error GoTo anula

    'If MsgBox("¿Desea anular el pedido " + vbCrLf + Me.lvwPedidos.SelectedItem.SubItems(3) + " ?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    If MsgBox("¿Desea anular el pedido " + vbCrLf + gNumeroPedido + " ?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then
    AnularPedido = False
    gNumeroPedido = ""
    Exit Function
End If
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "USP_PEDIDO_ANULA"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@idpedido", adInteger, adParamInput, , gNumeroPedido)
    oCmdEjec.Execute
    AnularPedido = True
    gNumeroPedido = ""
    Exit Function
anula:
    gNumeroPedido = ""
    AnularPedido = False

End Function
