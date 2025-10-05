Attribute VB_Name = "ModuloAlu"
' ===== ModuloALU =====
Option Explicit

Private Function ObtenerValorRegistro(nombreRegistro As String) As Long
    Select Case UCase(nombreRegistro)
        Case "EAX": ObtenerValorRegistro = EAX
        Case "EBX": ObtenerValorRegistro = EBX
        Case "ECX": ObtenerValorRegistro = ECX
        Case "EDX": ObtenerValorRegistro = EDX
        Case "ESI": ObtenerValorRegistro = ESI
        Case "EDI": ObtenerValorRegistro = EDI
        Case "EBP": ObtenerValorRegistro = EBP
        Case "ESP": ObtenerValorRegistro = ESP
        Case Else: ObtenerValorRegistro = 0
    End Select
End Function
Private Function ObtenerValorOperando(operando As String) As Long
    If operando = "" Then
        ObtenerValorOperando = 0
        Exit Function
    End If
    
    If EsNumero(operando) Then
        ObtenerValorOperando = ConvertirANumero(operando)
    Else
        ObtenerValorOperando = ObtenerValorRegistro(operando)
    End If
End Function
Private Sub ActualizarRegistro(nombreRegistro As String, valor As Long)
    Select Case UCase(nombreRegistro)
        Case "EAX": EAX = valor
        Case "EBX": EBX = valor
        Case "ECX": ECX = valor
        Case "EDX": EDX = valor
        Case "ESI": ESI = valor
        Case "EDI": EDI = valor
        Case "EBP": EBP = valor
        Case "ESP": ESP = valor
    End Select
End Sub
Private Sub ActualizarFlagsADD(operando1 As Long, operando2 As Long, resultado As Long)
    ZF = (resultado = 0)
    SF = (resultado And &H80000000) <> 0
    CF = (CLng(operando1) + CLng(operando2)) > &HFFFFFFFF
    OF = ((operando1 And &H80000000) = (operando2 And &H80000000)) And _
         ((operando1 And &H80000000) <> (resultado And &H80000000))
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    AF = ((operando1 And &HF) + (operando2 And &HF)) > &HF
End Sub
Private Sub ActualizarFlagsSUB(operando1 As Long, operando2 As Long, resultado As Long)
    ZF = (resultado = 0)
    SF = (resultado < 0)
    CF = (operando1 < operando2)
    OF = ((operando1 Xor operando2) And (operando1 Xor resultado) And &H80000000) <> 0
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    AF = ((operando1 And &HF) < (operando2 And &HF))
End Sub
Private Sub ActualizarFlagsMUL(resultado64 As Currency, resultado32 As Long)
    If resultado64 > 2147483647 Or resultado64 < -2147483648# Then
        CF = True
        OF = True
    Else
        CF = False
        OF = False
    End If
    ZF = (resultado32 = 0)
    SF = (resultado32 < 0)
    PF = (ContarBits1(resultado32 And &HFF) Mod 2 = 0)
End Sub
Private Sub ActualizarFlagsIMUL(operando1 As Long, operando2 As Long, resultado As Long)
    If resultado <> CLng(operando1) * CLng(operando2) Then
        CF = True
        OF = True
    Else
        CF = False
        OF = False
    End If
    ZF = (resultado = 0)
    SF = (resultado < 0)
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
End Sub
Private Function ContarBits1(valor As Long) As Integer
    Dim i As Integer, count As Integer
    count = 0
    For i = 0 To 31
        If (valor And (2 ^ i)) <> 0 Then count = count + 1
    Next i
    ContarBits1 = count
End Function
Private Function EsNumero(texto As String) As Boolean
    On Error GoTo ErrorHandler
    Dim temp As String
    temp = UCase(Trim(texto))
    If Left(temp, 2) = "0X" Then
        temp = "&H" & Mid(temp, 3)
    End If
    If InStr(temp, "&H") = 0 And Not IsNumeric(temp) Then
        EsNumero = False
        Exit Function
    End If
    Dim valor As Long
    valor = CLng(temp)
    EsNumero = True
    Exit Function
ErrorHandler:
    EsNumero = False
End Function
Private Function ConvertirANumero(texto As String) As Long
    Dim temp As String
    temp = UCase(Trim(texto))
    If Left(temp, 2) = "0X" Then
        temp = "&H" & Mid(temp, 3)
    End If
    ConvertirANumero = CLng(temp)
End Function
Public Sub EjecutarMOV(destino As String, origen As String)
    Dim valorOrigen As Long
    valorOrigen = ObtenerValorOperando(origen)
    ActualizarRegistro destino, valorOrigen
End Sub
Public Sub EjecutarADD(destino As String, origen As String)
    Dim valorDestino As Long, valorOrigen As Long, resultado As Long
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    resultado = valorDestino + valorOrigen
    ActualizarRegistro destino, resultado
    ActualizarFlagsADD valorDestino, valorOrigen, resultado
End Sub
Public Sub EjecutarSUB(destino As String, origen As String)
    Dim valorDestino As Long, valorOrigen As Long, resultado As Long
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    resultado = valorDestino - valorOrigen
    ActualizarRegistro destino, resultado
    ActualizarFlagsSUB valorDestino, valorOrigen, resultado
End Sub
Public Sub EjecutarMUL(operando As String)
    Dim valorOperando As Long, resultado As Long
    valorOperando = ObtenerValorOperando(operando)
    Dim resultado64 As Currency
    resultado64 = CDbl(EAX) * CDbl(valorOperando)
    resultado = EAX * valorOperando
    EAX = resultado
    EDX = 0
    ActualizarFlagsMUL resultado64, resultado
End Sub
Public Sub EjecutarDIV(operando As String)
    Dim valorOperando As Long
    valorOperando = ObtenerValorOperando(operando)
    If valorOperando = 0 Then
        MsgBox "ERROR: División por cero", vbCritical, "Error de Ejecución"
        Exit Sub
    End If
    Dim dividendo As Long
    dividendo = EAX
    Dim cociente As Long, residuo As Long
    cociente = dividendo \ valorOperando
    residuo = dividendo Mod valorOperando
    EAX = cociente
    EDX = residuo
    ZF = (cociente = 0)
    SF = (cociente < 0)
End Sub
Public Sub EjecutarIMUL(operando As String)
    Dim valorOperando As Long, resultado As Long
    valorOperando = ObtenerValorOperando(operando)
    resultado = CLng(EAX) * CLng(valorOperando)
    EAX = resultado
    EDX = 0
    ActualizarFlagsIMUL EAX, valorOperando, resultado
End Sub
Public Sub EjecutarIDIV(operando As String)
    Dim valorOperando As Long
    valorOperando = ObtenerValorOperando(operando)
    If valorOperando = 0 Then
        MsgBox "ERROR: División por cero", vbCritical, "Error de Ejecución"
        Exit Sub
    End If
    Dim dividendo As Long
    dividendo = EAX
    Dim cociente As Long, residuo As Long
    cociente = CLng(dividendo) \ CLng(valorOperando)
    residuo = CLng(dividendo) Mod CLng(valorOperando)
    EAX = cociente
    EDX = residuo
    ZF = (cociente = 0)
    SF = (cociente < 0)
End Sub
Public Sub EjecutarAND(destino As String, origen As String)
    Dim valorDestino As Long, valorOrigen As Long, resultado As Long
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    resultado = valorDestino And valorOrigen
    ActualizarRegistro destino, resultado
    ZF = (resultado = 0)
    SF = (resultado And &H80000000) <> 0
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    CF = False
    OF = False
    AF = False
End Sub
Public Sub EjecutarOR(destino As String, origen As String)
    Dim valorDestino As Long, valorOrigen As Long, resultado As Long
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    resultado = valorDestino Or valorOrigen
    ActualizarRegistro destino, resultado
    ZF = (resultado = 0)
    SF = (resultado And &H80000000) <> 0
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    CF = False
    OF = False
    AF = False
End Sub
Public Sub EjecutarXOR(destino As String, origen As String)
    Dim valorDestino As Long, valorOrigen As Long, resultado As Long
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    resultado = valorDestino Xor valorOrigen
    ActualizarRegistro destino, resultado
    ZF = (resultado = 0)
    SF = (resultado And &H80000000) <> 0
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    CF = False
    OF = False
    AF = False
End Sub
Public Sub EjecutarNOT(destino As String)
    Dim valor As Long
    valor = ObtenerValorRegistro(destino)
    valor = Not valor
    ActualizarRegistro destino, valor
End Sub
Public Sub EjecutarCMP(operando1 As String, operando2 As String)
    Dim val1 As Long, val2 As Long, resultado As Long
    val1 = ObtenerValorRegistro(operando1)
    val2 = ObtenerValorOperando(operando2)
    resultado = val1 - val2
    ActualizarFlagsSUB val1, val2, resultado
End Sub
Public Sub EjecutarTEST(operando1 As String, operando2 As String)
    Dim val1 As Long, val2 As Long, resultado As Long
    val1 = ObtenerValorRegistro(operando1)
    val2 = ObtenerValorOperando(operando2)
    resultado = val1 And val2
    ZF = (resultado = 0)
    SF = (resultado And &H80000000) <> 0
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    CF = False
    OF = False
    AF = False
End Sub
Public Sub EjecutarINC(destino As String)
    Dim valor As Long
    valor = ObtenerValorRegistro(destino)
    valor = valor + 1
    ActualizarRegistro destino, valor
    ZF = (valor = 0)
    SF = (valor < 0)
    PF = (ContarBits1(valor And &HFF) Mod 2 = 0)
End Sub
Public Sub EjecutarDEC(destino As String)
    Dim valor As Long
    valor = ObtenerValorRegistro(destino)
    valor = valor - 1
    ActualizarRegistro destino, valor
    ZF = (valor = 0)
    SF = (valor < 0)
    PF = (ContarBits1(valor And &HFF) Mod 2 = 0)
End Sub
Public Sub EjecutarSHL(destino As String, cantidad As String)
    Dim valor As Long, shift As Integer
    valor = ObtenerValorRegistro(destino)
    shift = CInt(ObtenerValorOperando(cantidad))
    valor = valor * (2 ^ shift)
    ActualizarRegistro destino, valor
    ZF = (valor = 0)
    SF = (valor And &H80000000) <> 0
    PF = (ContarBits1(valor And &HFF) Mod 2 = 0)
    CF = False
    OF = False
    AF = False
End Sub
Public Sub EjecutarSHR(destino As String, cantidad As String)
    Dim valor As Long, shift As Integer
    valor = ObtenerValorRegistro(destino)
    shift = CInt(ObtenerValorOperando(cantidad))
    valor = valor \ (2 ^ shift)
    ActualizarRegistro destino, valor
    ZF = (valor = 0)
    SF = (valor And &H80000000) <> 0
    PF = (ContarBits1(valor And &HFF) Mod 2 = 0)
    CF = False
    OF = False
    AF = False
End Sub
