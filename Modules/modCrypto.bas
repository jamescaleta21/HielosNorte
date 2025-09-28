Attribute VB_Name = "modCrypto"
' clsCrypto - clase que provee EncryptText / DecryptText (Base64 integrado)
Option Explicit

Private Const Base64Table As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

' --- Encripta texto con un desplazamiento fijo + Base64 ---
Public Function EncryptText(ByVal plainText As String) As String
    Dim i As Long
    Dim charCode As Long
    Dim result As String
    
    For i = 1 To Len(plainText)
        charCode = Asc(Mid$(plainText, i, 1))
        charCode = (charCode + 5) Mod 256
        result = result & Chr$(charCode)
    Next i
    
    EncryptText = Base64Encode(result)
End Function

' --- Desencripta texto producido por EncryptText ---
Public Function DecryptText(ByVal cipherText As String) As String
    Dim decoded As String
    Dim i As Long
    Dim charCode As Long
    Dim result As String
    
    decoded = Base64Decode(cipherText)
    
    For i = 1 To Len(decoded)
        charCode = Asc(Mid$(decoded, i, 1))
        charCode = (charCode - 5 + 256) Mod 256
        result = result & Chr$(charCode)
    Next i
    
    DecryptText = result
End Function

' -----------------------
' Base64 helpers (privados)
' -----------------------
Private Function Base64Encode(ByVal sText As String) As String
    Dim i As Long, n As Long
    Dim outStr As String
    Dim b1 As Long, b2 As Long, b3 As Long
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long
    
    outStr = ""
    If Len(sText) = 0 Then
        Base64Encode = ""
        Exit Function
    End If
    
    For i = 1 To Len(sText) Step 3
        b1 = Asc(Mid$(sText, i, 1))
        If i + 1 <= Len(sText) Then b2 = Asc(Mid$(sText, i + 1, 1)) Else b2 = 0
        If i + 2 <= Len(sText) Then b3 = Asc(Mid$(sText, i + 2, 1)) Else b3 = 0
        
        c1 = b1 \ 4
        c2 = ((b1 And 3) * 16) Or (b2 \ 16)
        c3 = ((b2 And 15) * 4) Or (b3 \ 64)
        c4 = b3 And 63
        
        outStr = outStr & Mid$(Base64Table, c1 + 1, 1)
        outStr = outStr & Mid$(Base64Table, c2 + 1, 1)
        
        If i + 1 <= Len(sText) Then
            outStr = outStr & Mid$(Base64Table, c3 + 1, 1)
        Else
            outStr = outStr & "="
        End If
        
        If i + 2 <= Len(sText) Then
            outStr = outStr & Mid$(Base64Table, c4 + 1, 1)
        Else
            outStr = outStr & "="
        End If
    Next i
    
    Base64Encode = outStr
End Function

' === REEMPLAZO COMPLETO DE Base64Decode EN clsCrypto ===
Private Function B64CharVal(ByVal ch As String) As Long
    Dim p As Long
    If ch = "=" Or Len(ch) = 0 Then
        B64CharVal = -1
        Exit Function
    End If
    p = InStr(1, Base64Table, ch, vbBinaryCompare)
    If p = 0 Then
        B64CharVal = -1
    Else
        B64CharVal = p - 1
    End If
End Function

Private Function Base64Decode(ByVal sText As String) As String
    Dim clean As String
    Dim L As Long, i As Long
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long
    Dim out() As Byte, idx As Long

    ' Limpieza básica
    clean = Replace(sText, vbCr, "")
    clean = Replace(clean, vbLf, "")
    clean = Replace(clean, " ", "")

    L = Len(clean)
    If L = 0 Then
        Base64Decode = ""
        Exit Function
    End If

    ' Tamaño máximo posible (luego hacemos Preserve)
    ReDim out(((L + 3) \ 4) * 3 - 1)
    idx = 0

    For i = 1 To L Step 4
        c1 = B64CharVal(Mid$(clean, i, 1))
        c2 = B64CharVal(Mid$(clean, i + 1, 1))
        c3 = B64CharVal(Mid$(clean, i + 2, 1))
        c4 = B64CharVal(Mid$(clean, i + 3, 1))

        ' c1 y c2 deben existir siempre
        If c1 < 0 Or c2 < 0 Then Exit For

        ' Primer byte
        out(idx) = ((c1 * 4) Or (c2 \ 16)) And &HFF
        idx = idx + 1

        ' Segundo byte (si no hay '=' en la 3ra pos)
        If c3 >= 0 Then
            out(idx) = (((c2 And &HF) * 16) Or (c3 \ 4)) And &HFF
            idx = idx + 1

            ' Tercer byte (si no hay '=' en la 4ta pos)
            If c4 >= 0 Then
                out(idx) = (((c3 And 3) * 64) Or c4) And &HFF
                idx = idx + 1
            End If
        End If
    Next i

    If idx = 0 Then
        Base64Decode = ""
    Else
        ReDim Preserve out(idx - 1)
        ' Convertimos los bytes a String (1 byte -> 1 char)
        Base64Decode = StrConv(out, vbUnicode)
    End If
End Function

