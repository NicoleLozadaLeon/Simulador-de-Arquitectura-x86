Attribute VB_Name = "ModuloAlu"
' ===== ModuloALU =====
Option Explicit

' Función para obtener valor de un registro por nombre
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

' Función para obtener valor de un operando (puede ser registro o número)
Private Function ObtenerValorOperando(operando As String) As Long
    On Error GoTo ErrorHandler
    ' Si es un número (hexadecimal o decimal)
    If EsNumero(operando) Then
        ObtenerValorOperando = ConvertirANumero(operando)
    Else
        ' Si es un registro
        ObtenerValorOperando = ObtenerValorRegistro(operando)
    End If
    Exit Function
ErrorHandler:
    ObtenerValorOperando = 0
End Function

' Función para actualizar un registro por nombre
Private Sub ActualizarRegistro(nombreRegistro As String, valor As Long)
    On Error GoTo ErrorHandler
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
    Exit Sub
ErrorHandler:
    ' En caso de overflow, establecer valor máximo/mínimo
    If valor > 2147483647 Then
        valor = 2147483647
    ElseIf valor < -2147483648# Then
        valor = -2147483648#
    End If
    Resume
End Sub

' Función para actualizar flags después de ADD - CORREGIDA
Private Sub ActualizarFlagsADD(operando1 As Long, operando2 As Long, resultado As Long)
    On Error GoTo ErrorHandler
    ZF = (resultado = 0)
    SF = (resultado < 0)
    ' Cálculo seguro de CF
    If operando1 > 0 And operando2 > 0 And resultado < 0 Then
        CF = True
    ElseIf operando1 < 0 And operando2 < 0 And resultado >= 0 Then
        CF = True
    Else
        CF = False
    End If
    ' Cálculo seguro de OF
    If ((operando1 Xor operando2) And &H80000000) = 0 Then
        If ((operando1 Xor resultado) And &H80000000) <> 0 Then
            OF = True
        Else
            OF = False
        End If
    Else
        OF = False
    End If
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    AF = ((operando1 And &HF) + (operando2 And &HF)) > &HF
    Exit Sub
ErrorHandler:
    ' En caso de error, establecer flags por defecto
    ZF = False: SF = False: CF = True: OF = True: PF = False: AF = False
End Sub

' Actualizar flags para operación SUB - CORREGIDA
Private Sub ActualizarFlagsSUB(operando1 As Long, operando2 As Long, resultado As Long)
    On Error GoTo ErrorHandler
    ZF = (resultado = 0)
    SF = (resultado < 0)
    ' Cálculo seguro de CF
    CF = (operando1 < operando2)
    ' Cálculo seguro de OF
    OF = ((operando1 And &H80000000) <> (operando2 And &H80000000)) And _
         ((operando1 And &H80000000) <> (resultado And &H80000000))
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    AF = ((operando1 And &HF) < (operando2 And &HF))
    Exit Sub
ErrorHandler:
    ZF = False: SF = False: CF = True: OF = True: PF = False: AF = False
End Sub

' Actualizar flags para operación MUL - CORREGIDA
Private Sub ActualizarFlagsMUL(resultado64 As Currency, resultado32 As Long)
    On Error GoTo ErrorHandler
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
    Exit Sub
ErrorHandler:
    CF = True: OF = True: ZF = False: SF = False: PF = False
End Sub

' Actualizar flags para operación IMUL - CORREGIDA
Private Sub ActualizarFlagsIMUL(operando1 As Long, operando2 As Long, resultado As Long)
    On Error GoTo ErrorHandler
    ' Verificar overflow de manera segura
    Dim temp As Double
    temp = CDbl(operando1) * CDbl(operando2)
    If temp > 2147483647 Or temp < -2147483648# Then
        CF = True
        OF = True
    Else
        CF = False
        OF = False
    End If
    ZF = (resultado = 0)
    SF = (resultado < 0)
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    Exit Sub
ErrorHandler:
    CF = True: OF = True: ZF = False: SF = False: PF = False
End Sub

' ========== FUNCIONES AUXILIARES EXISTENTES ==========

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
    If left(temp, 2) = "0X" Then
        temp = "&H" & Mid(temp, 3)
    End If
    If InStr(temp, "&H") = 0 And Not IsNumeric(temp) Then
        EsNumero = False
        Exit Function
    End If
    Dim valor As Long
    valor = CLng(val(temp)) ' Usar Val() para mayor seguridad
    EsNumero = True
    Exit Function
ErrorHandler:
    EsNumero = False
End Function

Private Function ConvertirANumero(texto As String) As Long
    On Error GoTo ErrorHandler
    Dim temp As String
    temp = UCase(Trim(texto))
    If left(temp, 2) = "0X" Then
        temp = "&H" & Mid(temp, 3)
    End If
    ConvertirANumero = CLng(val(temp)) ' Usar Val() para mayor seguridad
    Exit Function
ErrorHandler:
    ConvertirANumero = 0
End Function

' ========== OPERACIONES ARITMÉTICAS CORREGIDAS ==========

Public Sub EjecutarMOV(destino As String, origen As String)
    On Error GoTo ErrorHandler
    Dim valorOrigen As Long
    valorOrigen = ObtenerValorOperando(origen)
    ActualizarRegistro destino, valorOrigen
    ' MOV no afecta flags
    Exit Sub
ErrorHandler:
    ' En caso de error, establecer valor 0
    ActualizarRegistro destino, 0
End Sub

Public Sub EjecutarADD(destino As String, origen As String)
    On Error GoTo ErrorHandler
    Dim valorDestino As Long, valorOrigen As Long, resultado As Long
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    
    ' Cálculo seguro para evitar overflow
    resultado = CLng(CDbl(valorDestino) + CDbl(valorOrigen))
    
    ActualizarRegistro destino, resultado
    ActualizarFlagsADD valorDestino, valorOrigen, resultado
    Exit Sub
ErrorHandler:
    ' En caso de overflow, establecer valor máximo/mínimo
    If CDbl(valorDestino) + CDbl(valorOrigen) > 2147483647 Then
        ActualizarRegistro destino, 2147483647
    ElseIf CDbl(valorDestino) + CDbl(valorOrigen) < -2147483648# Then
        ActualizarRegistro destino, -2147483648#
    Else
        ActualizarRegistro destino, 0
    End If
    ActualizarFlagsADD valorDestino, valorOrigen, ObtenerValorRegistro(destino)
End Sub

Public Sub EjecutarSUB(destino As String, origen As String)
    On Error GoTo ErrorHandler
    Dim valorDestino As Long, valorOrigen As Long, resultado As Long
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    
    ' Cálculo seguro para evitar overflow
    resultado = CLng(CDbl(valorDestino) - CDbl(valorOrigen))
    
    ActualizarRegistro destino, resultado
    ActualizarFlagsSUB valorDestino, valorOrigen, resultado
    Exit Sub
ErrorHandler:
    ' En caso de overflow, establecer valor máximo/mínimo
    If CDbl(valorDestino) - CDbl(valorOrigen) > 2147483647 Then
        ActualizarRegistro destino, 2147483647
    ElseIf CDbl(valorDestino) - CDbl(valorOrigen) < -2147483648# Then
        ActualizarRegistro destino, -2147483648#
    Else
        ActualizarRegistro destino, 0
    End If
    ActualizarFlagsSUB valorDestino, valorOrigen, ObtenerValorRegistro(destino)
End Sub

Public Sub EjecutarMUL(operando As String)
    On Error GoTo ErrorHandler
    Dim valorOperando As Long
    valorOperando = ObtenerValorOperando(operando)
    
    ' Usar Double para cálculos grandes
    Dim resultado64 As Double
    resultado64 = CDbl(EAX) * CDbl(valorOperando)
    
    ' Verificar límites antes de asignar
    If resultado64 > 2147483647 Then
        EAX = 2147483647
        CF = True: OF = True
    ElseIf resultado64 < -2147483648# Then
        EAX = -2147483648#
        CF = True: OF = True
    Else
        EAX = CLng(resultado64)
        CF = False: OF = False
    End If
    
    EDX = 0 ' En esta implementación simplificada
    ActualizarFlagsMUL resultado64, EAX
    Exit Sub
ErrorHandler:
    EAX = 0: EDX = 0
    CF = True: OF = True: ZF = False: SF = False: PF = False
End Sub

Public Sub EjecutarDIV(operando As String)
    On Error GoTo ErrorHandler
    Dim valorOperando As Long
    valorOperando = ObtenerValorOperando(operando)
    
    If valorOperando = 0 Then
        MsgBox "ERROR: División por cero", vbCritical, "Error de Ejecución"
        Exit Sub
    End If
    
    ' Usar Double para evitar overflow
    Dim dividendo As Double
    dividendo = CDbl(EAX)
    
    Dim cociente As Long, residuo As Long
    cociente = CLng(dividendo / CDbl(valorOperando))
    residuo = CLng(dividendo Mod CDbl(valorOperando))
    
    ' Verificar límites
    If cociente > 2147483647 Then cociente = 2147483647
    If cociente < -2147483648# Then cociente = -2147483648#
    If residuo > 2147483647 Then residuo = 2147483647
    If residuo < -2147483648# Then residuo = -2147483648#
    
    EAX = cociente
    EDX = residuo
    ZF = (cociente = 0)
    SF = (cociente < 0)
    Exit Sub
ErrorHandler:
    EAX = 0: EDX = 0
    ZF = False: SF = False
    MsgBox "Error en división: " & Err.Description, vbExclamation
End Sub

Public Sub EjecutarIMUL(operando As String)
    On Error GoTo ErrorHandler
    Dim valorOperando As Long
    valorOperando = ObtenerValorOperando(operando)
    
    ' Usar Double para cálculos seguros
    Dim resultado As Double
    resultado = CDbl(EAX) * CDbl(valorOperando)
    
    ' Verificar límites
    If resultado > 2147483647 Then
        EAX = 2147483647
        CF = True: OF = True
    ElseIf resultado < -2147483648# Then
        EAX = -2147483648#
        CF = True: OF = True
    Else
        EAX = CLng(resultado)
        CF = False: OF = False
    End If
    
    EDX = 0
    ActualizarFlagsIMUL EAX, valorOperando, EAX
    Exit Sub
ErrorHandler:
    EAX = 0: EDX = 0
    CF = True: OF = True: ZF = False: SF = False: PF = False
End Sub

Public Sub EjecutarIDIV(operando As String)
    On Error GoTo ErrorHandler
    Dim valorOperando As Long
    valorOperando = ObtenerValorOperando(operando)
    
    If valorOperando = 0 Then
        MsgBox "ERROR: División por cero", vbCritical, "Error de Ejecución"
        Exit Sub
    End If
    
    ' Usar Double para cálculos seguros
    Dim dividendo As Double
    dividendo = CDbl(EAX)
    
    Dim cociente As Long, residuo As Long
    cociente = CLng(dividendo / CDbl(valorOperando))
    residuo = CLng(dividendo Mod CDbl(valorOperando))
    
    ' Verificar límites
    If cociente > 2147483647 Then cociente = 2147483647
    If cociente < -2147483648# Then cociente = -2147483648#
    If residuo > 2147483647 Then residuo = 2147483647
    If residuo < -2147483648# Then residuo = -2147483648#
    
    EAX = cociente
    EDX = residuo
    ZF = (cociente = 0)
    SF = (cociente < 0)
    Exit Sub
ErrorHandler:
    EAX = 0: EDX = 0
    ZF = False: SF = False
    MsgBox "Error en división con signo: " & Err.Description, vbExclamation
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
