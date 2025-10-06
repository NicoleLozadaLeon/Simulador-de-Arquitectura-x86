Attribute VB_Name = "ModuloGrafico"
' ===== MODULO GRAFICO CPU =====
' Todas las funciones tienen nombres únicos para evitar conflictos
Option Explicit

' Variable para llevar el control de la próxima entrada disponible
Dim ProximaEntradaDisponible As Integer

Sub EjecutarInstruccionGrafico()
    Dim ni As Integer 'numero instruccion
    Dim instruccionCompleta As String
    Dim Opcode As String
    Dim operando1 As String
    Dim operando2 As String
    Dim i As Integer
    
    'Inicializar próxima entrada disponible
    ProximaEntradaDisponible = 0
    
    QuitarColoresGrafico
    
    'Ir a instrucción - CONTADOR en C31
    Range("C31").Select
    ni = ActiveCell.FormulaR1C1  'tomo el numero de instrucción
    
    If ni > 20 Then
        MsgBox "Fin del programa"
        GoTo Error
    Else
        Range("C8").Select 'anterior a la primera instrucción
        For i = 0 To ni
            ActiveCell.offset(1, 0).Range("A1").Select
        Next
        
        'Resaltar instrucción actual (C9:C29)
        ActiveCell.Range("A1:A1").Select ' Solo la celda de instrucción
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        'Obtener instrucción completa
        instruccionCompleta = Trim(ActiveCell.FormulaR1C1)
        
        'Si está vacío, fin del programa
        If instruccionCompleta = "" Then
            MsgBox "Fin del programa"
            GoTo Error
        End If
        
        'Parsear instrucción ensamblador
        Call ParsearInstruccionEnsambladorGrafico(instruccionCompleta, Opcode, operando1, operando2)
        
        'Ejecutar instrucción según opcode
        Select Case UCase(Opcode)
            Case "MOV"
                EjecutarMOVGrafico operando1, operando2
            Case "ADD"
                EjecutarADDGrafico operando1, operando2
            Case "SUB"
                EjecutarSUBGrafico operando1, operando2
            Case "INC"
                EjecutarINCGrafico operando1
            Case "DEC"
                EjecutarDECGrafico operando1
            Case "MUL"
                EjecutarMULGrafico operando1
            Case "DIV"
                EjecutarDIVGrafico operando1
            Case "JMP"
                EjecutarJMPGrafico operando1
            Case "JZ", "JE"
                EjecutarJZGrafico operando1
            Case "JNZ", "JNE"
                EjecutarJNZGrafico operando1
            Case "CMP"
                EjecutarCMPGrafico operando1, operando2
            Case "NOP"
                'No operation - no hace nada
            Case Else
                MsgBox "Instrucción no reconocida: " & Opcode
                GoTo Error
        End Select
        
        'Aumentar contador de instrucción en C31 (a menos que sea salto)
        If UCase(Opcode) <> "JMP" And UCase(Opcode) <> "JZ" And UCase(Opcode) <> "JNZ" And _
           UCase(Opcode) <> "JE" And UCase(Opcode) <> "JNE" Then
            ni = ni + 1
            Range("C31").Select
            ActiveCell.FormulaR1C1 = ni
        End If
    End If
    
    Exit Sub
    
Error:
End Sub

Sub ParsearInstruccionEnsambladorGrafico(instruccion As String, Opcode As String, operando1 As String, operando2 As String)
    Dim partes() As String
    Dim i As Integer
    
    'Limpiar variables
    Opcode = ""
    operando1 = ""
    operando2 = ""
    
    'Dividir por espacios
    partes = Split(instruccion, " ")
    
    'Obtener opcode (primera palabra)
    If UBound(partes) >= 0 Then
        Opcode = UCase(Trim(partes(0)))
    End If
    
    'Obtener operandos
    If UBound(partes) >= 1 Then
        'Unir todos los operandos y luego separar por coma
        Dim todosOperandos As String
        todosOperandos = Trim(Mid(instruccion, Len(Opcode) + 1))
        
        'Dividir por coma
        Dim operandos() As String
        operandos = Split(todosOperandos, ",")
        
        If UBound(operandos) >= 0 Then
            operando1 = Trim(operandos(0))
        End If
        
        If UBound(operandos) >= 1 Then
            operando2 = Trim(operandos(1))
        End If
    End If
End Sub

' ========== INSTRUCCIONES ENSAMBLADOR x86 CORREGIDAS ==========

Sub EjecutarMOVGrafico(destino As String, origen As String)
    Dim valor As Integer
    
    'Obtener valor del origen
    If EsRegistroGrafico(origen) Then
        valor = ObtenerValorRegistroGrafico(origen)
    ElseIf EsNumeroGrafico(origen) Then
        valor = CInt(origen)
        'Colocar valor numérico en la próxima entrada disponible
        ColocarValorEnEntrada valor
    Else
        MsgBox "Error: Operando inválido en MOV"
        Exit Sub
    End If
    
    'Mover a destino
    If EsRegistroGrafico(destino) Then
        AsignarValorRegistroGrafico destino, valor
        'Si el destino es un registro, también mostrar en salida
        If destino = "EAX" Or destino = "ACUMULADOR" Then
            ColocarValorEnSalida valor
        End If
    Else
        MsgBox "Error: Destino inválido en MOV"
    End If
End Sub

Sub EjecutarADDGrafico(destino As String, origen As String)
    Dim valorDestino As Integer
    Dim valorOrigen As Integer
    Dim resultado As Integer
    
    'Obtener valores
    If EsRegistroGrafico(destino) Then
        valorDestino = ObtenerValorRegistroGrafico(destino)
    Else
        MsgBox "Error: Destino inválido en ADD"
        Exit Sub
    End If
    
    If EsRegistroGrafico(origen) Then
        valorOrigen = ObtenerValorRegistroGrafico(origen)
    ElseIf EsNumeroGrafico(origen) Then
        valorOrigen = CInt(origen)
        ColocarValorEnEntrada valorOrigen
    Else
        MsgBox "Error: Operando inválido en ADD"
        Exit Sub
    End If
    
    'Realizar suma
    resultado = valorDestino + valorOrigen
    
    'Mostrar operación en entradas
    ColocarValorEnEntrada valorDestino
    ColocarValorEnEntrada valorOrigen
    
    'Asignar resultado
    AsignarValorRegistroGrafico destino, resultado
    
    'Mostrar resultado en salida
    ColocarValorEnSalida resultado
    
    'Actualizar flags
    ActualizarFlagsADDGrafico valorDestino, valorOrigen, resultado
End Sub

Sub EjecutarSUBGrafico(destino As String, origen As String)
    Dim valorDestino As Integer
    Dim valorOrigen As Integer
    Dim resultado As Integer
    
    'Obtener valores
    If EsRegistroGrafico(destino) Then
        valorDestino = ObtenerValorRegistroGrafico(destino)
    Else
        MsgBox "Error: Destino inválido en SUB"
        Exit Sub
    End If
    
    If EsRegistroGrafico(origen) Then
        valorOrigen = ObtenerValorRegistroGrafico(origen)
    ElseIf EsNumeroGrafico(origen) Then
        valorOrigen = CInt(origen)
        ColocarValorEnEntrada valorOrigen
    Else
        MsgBox "Error: Operando inválido en SUB"
        Exit Sub
    End If
    
    'Mostrar operación en entradas
    ColocarValorEnEntrada valorDestino
    ColocarValorEnEntrada valorOrigen
    
    'Realizar resta
    resultado = valorDestino - valorOrigen
    
    'Asignar resultado
    AsignarValorRegistroGrafico destino, resultado
    
    'Mostrar resultado en salida
    ColocarValorEnSalida resultado
    
    'Actualizar flags
    ActualizarFlagsSUBGrafico valorDestino, valorOrigen, resultado
End Sub

Sub EjecutarINCGrafico(registro As String)
    Dim valor As Integer
    
    If EsRegistroGrafico(registro) Then
        valor = ObtenerValorRegistroGrafico(registro)
        'Mostrar valor original en entrada
        ColocarValorEnEntrada valor
        ColocarValorEnEntrada 1 'Para mostrar el incremento
        
        valor = valor + 1
        AsignarValorRegistroGrafico registro, valor
        
        'Mostrar resultado en salida
        ColocarValorEnSalida valor
    Else
        MsgBox "Error: Registro inválido en INC"
    End If
End Sub

Sub EjecutarDECGrafico(registro As String)
    Dim valor As Integer
    
    If EsRegistroGrafico(registro) Then
        valor = ObtenerValorRegistroGrafico(registro)
        'Mostrar valor original en entrada
        ColocarValorEnEntrada valor
        ColocarValorEnEntrada 1 'Para mostrar el decremento
        
        valor = valor - 1
        AsignarValorRegistroGrafico registro, valor
        
        'Mostrar resultado en salida
        ColocarValorEnSalida valor
    Else
        MsgBox "Error: Registro inválido en DEC"
    End If
End Sub

Sub EjecutarMULGrafico(operando As String)
    Dim valorEAX As Integer
    Dim valorOperando As Integer
    Dim resultado As Integer
    
    'Obtener valor de EAX (Acumulador)
    valorEAX = ObtenerValorRegistroGrafico("EAX")
    
    'Obtener valor del operando
    If EsRegistroGrafico(operando) Then
        valorOperando = ObtenerValorRegistroGrafico(operando)
    ElseIf EsNumeroGrafico(operando) Then
        valorOperando = CInt(operando)
        ColocarValorEnEntrada valorOperando
    Else
        MsgBox "Error: Operando inválido en MUL"
        Exit Sub
    End If
    
    'Mostrar operación en entradas
    ColocarValorEnEntrada valorEAX
    ColocarValorEnEntrada valorOperando
    
    'Realizar multiplicación
    resultado = valorEAX * valorOperando
    
    'Asignar resultado a EAX
    AsignarValorRegistroGrafico "EAX", resultado
    
    'Mostrar resultado en salida
    ColocarValorEnSalida resultado
    
    'Actualizar flags
    ActualizarFlagsMULGrafico valorEAX, valorOperando, resultado
End Sub

Sub EjecutarDIVGrafico(operando As String)
    Dim valorEAX As Integer
    Dim valorOperando As Integer
    Dim cociente As Integer
    Dim residuo As Integer
    
    'Obtener valor de EAX (Acumulador)
    valorEAX = ObtenerValorRegistroGrafico("EAX")
    
    'Obtener valor del operando
    If EsRegistroGrafico(operando) Then
        valorOperando = ObtenerValorRegistroGrafico(operando)
    ElseIf EsNumeroGrafico(operando) Then
        valorOperando = CInt(operando)
        ColocarValorEnEntrada valorOperando
    Else
        MsgBox "Error: Operando inválido en DIV"
        Exit Sub
    End If
    
    'Verificar división por cero
    If valorOperando = 0 Then
        MsgBox "Error: División por cero"
        Exit Sub
    End If
    
    'Mostrar operación en entradas
    ColocarValorEnEntrada valorEAX
    ColocarValorEnEntrada valorOperando
    
    'Realizar división
    cociente = valorEAX \ valorOperando
    residuo = valorEAX Mod valorOperando
    
    'Asignar cociente a EAX y residuo a EDX
    AsignarValorRegistroGrafico "EAX", cociente
    AsignarValorRegistroGrafico "EDX", residuo
    
    'Mostrar resultado en salida
    ColocarValorEnSalida cociente
    
    'Si hay residuo, mostrarlo también
    If residuo > 0 Then
        ColocarValorEnSalida residuo
    End If
End Sub

Sub EjecutarJMPGrafico(direccion As String)
    Dim nuevaDireccion As Integer
    
    If EsNumeroGrafico(direccion) Then
        nuevaDireccion = CInt(direccion)
        Range("C31").Select 'Contador en C31
        ActiveCell.FormulaR1C1 = nuevaDireccion
    Else
        MsgBox "Error: Dirección inválida en JMP"
    End If
End Sub

Sub EjecutarJZGrafico(direccion As String)
    'Saltar si Zero Flag está activo
    If ObtenerFlagGrafico("ZF") = 1 Then
        EjecutarJMPGrafico direccion
    Else
        'Incrementar contador normalmente
        Dim ni As Integer
        Range("C31").Select 'Contador en C31
        ni = ActiveCell.FormulaR1C1
        ni = ni + 1
        ActiveCell.FormulaR1C1 = ni
    End If
End Sub

Sub EjecutarJNZGrafico(direccion As String)
    'Saltar si Zero Flag NO está activo
    If ObtenerFlagGrafico("ZF") = 0 Then
        EjecutarJMPGrafico direccion
    Else
        'Incrementar contador normalmente
        Dim ni As Integer
        Range("C31").Select 'Contador en C31
        ni = ActiveCell.FormulaR1C1
        ni = ni + 1
        ActiveCell.FormulaR1C1 = ni
    End If
End Sub

Sub EjecutarCMPGrafico(operando1 As String, operando2 As String)
    Dim valor1 As Integer
    Dim valor2 As Integer
    Dim resultado As Integer
    
    'Obtener valores
    If EsRegistroGrafico(operando1) Then
        valor1 = ObtenerValorRegistroGrafico(operando1)
    ElseIf EsNumeroGrafico(operando1) Then
        valor1 = CInt(operando1)
        ColocarValorEnEntrada valor1
    Else
        MsgBox "Error: Operando 1 inválido en CMP"
        Exit Sub
    End If
    
    If EsRegistroGrafico(operando2) Then
        valor2 = ObtenerValorRegistroGrafico(operando2)
    ElseIf EsNumeroGrafico(operando2) Then
        valor2 = CInt(operando2)
        ColocarValorEnEntrada valor2
    Else
        MsgBox "Error: Operando 2 inválido en CMP"
        Exit Sub
    End If
    
    'Realizar comparación (resta sin guardar resultado)
    resultado = valor1 - valor2
    
    'Actualizar flags basado en la comparación
    ActualizarFlagsCMPGrafico valor1, valor2, resultado
End Sub

' ========== FUNCIONES AUXILIARES MEJORADAS ==========

Sub ColocarValorEnEntrada(valor As Integer)
    'Colocar valor en la próxima entrada disponible (F27:F30)
    If ProximaEntradaDisponible < 4 Then
        Select Case ProximaEntradaDisponible
            Case 0
                Range("F27").Select 'Entrada0
            Case 1
                Range("F28").Select 'Entrada1
            Case 2
                Range("F29").Select 'Entrada2
            Case 3
                Range("F30").Select 'Entrada3
        End Select
        ActiveCell.FormulaR1C1 = CStr(valor)
        ColorCeldaActivaGrafico
        ProximaEntradaDisponible = ProximaEntradaDisponible + 1
    End If
End Sub

Sub ColocarValorEnSalida(valor As Integer)
    'Buscar próxima salida disponible (H27:H30)
    Dim i As Integer
    For i = 0 To 3
        Select Case i
            Case 0
                Range("H27").Select 'Salida0
            Case 1
                Range("H28").Select 'Salida1
            Case 2
                Range("H29").Select 'Salida2
            Case 3
                Range("H30").Select 'Salida3
        End Select
        If ActiveCell.FormulaR1C1 = "0" Or ActiveCell.FormulaR1C1 = "" Then
            ActiveCell.FormulaR1C1 = CStr(valor)
            ColorCeldaActivaGrafico
            Exit For
        End If
    Next i
End Sub

Function EsRegistroGrafico(nombre As String) As Boolean
    Dim registro As String
    registro = UCase(Trim(nombre))
    
    Select Case registro
        Case "EAX", "EBX", "ECX", "EDX", "ESI", "EDI", "EBP", "ESP", _
             "AX", "BX", "CX", "DX", "SI", "DI", "BP", "SP", _
             "AL", "BL", "CL", "DL", "AH", "BH", "CH", "DH", _
             "ACUMULADOR", "REGISTRO1", "REGISTRO2", "CONTADOR", "ESTADO"
            EsRegistroGrafico = True
        Case Else
            EsRegistroGrafico = False
    End Select
End Function

Function EsNumeroGrafico(texto As String) As Boolean
    On Error GoTo ErrorHandler
    Dim temp As String
    temp = Trim(texto)
    
    'Verificar si es número (incluyendo negativos)
    If IsNumeric(temp) Then
        EsNumeroGrafico = True
    Else
        EsNumeroGrafico = False
    End If
    Exit Function
ErrorHandler:
    EsNumeroGrafico = False
End Function

Function ObtenerValorRegistroGrafico(nombre As String) As Integer
    Dim registro As String
    registro = UCase(Trim(nombre))
    
    Select Case registro
        Case "EAX", "AX", "ACUMULADOR"
            Range("M11").Select 'ACUMULADOR en M11
            ObtenerValorRegistroGrafico = CInt(ActiveCell.FormulaR1C1)
        Case "EBX", "BX", "REGISTRO1"
            Range("L23").Select 'REGISTRO1 en L23
            ObtenerValorRegistroGrafico = CInt(ActiveCell.FormulaR1C1)
        Case "ECX", "CX", "REGISTRO2"
            Range("N23").Select 'REGISTRO2 en N23
            ObtenerValorRegistroGrafico = CInt(ActiveCell.FormulaR1C1)
        Case "EDX", "DX", "CONTADOR"
            Range("C31").Select 'CONTADOR en C31
            ObtenerValorRegistroGrafico = CInt(ActiveCell.FormulaR1C1)
        Case "ESTADO"
            Range("P16").Select 'ESTADO en P16
            ObtenerValorRegistroGrafico = CInt(ActiveCell.FormulaR1C1)
        Case Else
            ObtenerValorRegistroGrafico = 0
    End Select
End Function

Sub AsignarValorRegistroGrafico(nombre As String, valor As Integer)
    Dim registro As String
    registro = UCase(Trim(nombre))
    
    Select Case registro
        Case "EAX", "AX", "ACUMULADOR"
            Range("M11").Select 'ACUMULADOR en M11
            ActiveCell.FormulaR1C1 = CStr(valor)
            ColorCeldaActivaGrafico
        Case "EBX", "BX", "REGISTRO1"
            Range("L23").Select 'REGISTRO1 en L23
            ActiveCell.FormulaR1C1 = CStr(valor)
            ColorCeldaActivaGrafico
        Case "ECX", "CX", "REGISTRO2"
            Range("N23").Select 'REGISTRO2 en N23
            ActiveCell.FormulaR1C1 = CStr(valor)
            ColorCeldaActivaGrafico
        Case "EDX", "DX", "CONTADOR"
            Range("C31").Select 'CONTADOR en C31
            ActiveCell.FormulaR1C1 = CStr(valor)
            ColorCeldaActivaGrafico
        Case "ESTADO"
            Range("P16").Select 'ESTADO en P16
            ActiveCell.FormulaR1C1 = CStr(valor)
            ColorCeldaActivaGrafico
    End Select
End Sub

Function ObtenerFlagGrafico(nombreFlag As String) As Integer
    'En esta implementación simplificada, usamos el registro ESTADO
    'Bit 0: Zero Flag (ZF)
    'Bit 1: Sign Flag (SF)
    'Bit 2: Carry Flag (CF)
    'Bit 3: Overflow Flag (OF)
    
    Dim estado As Integer
    estado = ObtenerValorRegistroGrafico("ESTADO")
    
    Select Case UCase(nombreFlag)
        Case "ZF" 'Zero Flag - bit 0
            ObtenerFlagGrafico = estado And 1
        Case "SF" 'Sign Flag - bit 1
            ObtenerFlagGrafico = (estado And 2) \ 2
        Case "CF" 'Carry Flag - bit 2
            ObtenerFlagGrafico = (estado And 4) \ 4
        Case "OF" 'Overflow Flag - bit 3
            ObtenerFlagGrafico = (estado And 8) \ 8
        Case Else
            ObtenerFlagGrafico = 0
    End Select
End Function

Sub ActualizarFlagsADDGrafico(valor1 As Integer, valor2 As Integer, resultado As Integer)
    Dim estado As Integer
    estado = 0
    
    'Zero Flag (resultado = 0)
    If resultado = 0 Then estado = estado Or 1
    
    'Sign Flag (resultado < 0)
    If resultado < 0 Then estado = estado Or 2
    
    'Carry Flag (overflow sin signo)
    If (valor1 > 0 And valor2 > 0 And resultado < 0) Or _
       (valor1 < 0 And valor2 < 0 And resultado >= 0) Then
        estado = estado Or 4
    End If
    
    'Overflow Flag (overflow con signo)
    If ((valor1 Xor valor2) And &H8000) = 0 Then
        If ((valor1 Xor resultado) And &H8000) <> 0 Then
            estado = estado Or 8
        End If
    End If
    
    AsignarValorRegistroGrafico "ESTADO", estado
End Sub

Sub ActualizarFlagsSUBGrafico(valor1 As Integer, valor2 As Integer, resultado As Integer)
    Dim estado As Integer
    estado = 0
    
    'Zero Flag (resultado = 0)
    If resultado = 0 Then estado = estado Or 1
    
    'Sign Flag (resultado < 0)
    If resultado < 0 Then estado = estado Or 2
    
    'Carry Flag (préstamo en resta sin signo)
    If valor1 < valor2 Then estado = estado Or 4
    
    'Overflow Flag (overflow con signo)
    If ((valor1 Xor valor2) And &H8000) <> 0 Then
        If ((valor1 Xor resultado) And &H8000) <> 0 Then
            estado = estado Or 8
        End If
    End If
    
    AsignarValorRegistroGrafico "ESTADO", estado
End Sub

Sub ActualizarFlagsCMPGrafico(valor1 As Integer, valor2 As Integer, resultado As Integer)
    'CMP es como SUB pero no guarda el resultado
    ActualizarFlagsSUBGrafico valor1, valor2, resultado
End Sub

Sub ActualizarFlagsMULGrafico(valor1 As Integer, valor2 As Integer, resultado As Integer)
    Dim estado As Integer
    estado = 0
    
    'Zero Flag (resultado = 0)
    If resultado = 0 Then estado = estado Or 1
    
    'Sign Flag (resultado < 0)
    If resultado < 0 Then estado = estado Or 2
    
    'Carry Flag y Overflow Flag (si hay overflow)
    If resultado > 32767 Or resultado < -32768 Then
        estado = estado Or 4 'Carry Flag
        estado = estado Or 8 'Overflow Flag
    End If
    
    AsignarValorRegistroGrafico "ESTADO", estado
End Sub

' ========== FUNCIONES ORIGINALES DEL GRÁFICO ==========

Sub ResetearGrafico()
    QuitarColoresGrafico

    'SALIDAS en H27:H30
    Range("H27").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("H28").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("H29").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("H30").Select
    ActiveCell.FormulaR1C1 = "0"
    
    'REGISTROS
    Range("N23").Select 'REGISTRO2 en N23
    ActiveCell.FormulaR1C1 = "0"
    Range("L23").Select 'REGISTRO1 en L23
    ActiveCell.FormulaR1C1 = "0"
    Range("C31").Select 'CONTADOR en C31
    ActiveCell.FormulaR1C1 = "0"
    Range("M11").Select 'ACUMULADOR en M11
    ActiveCell.FormulaR1C1 = "0"
    Range("P16").Select 'ESTADO en P16
    ActiveCell.FormulaR1C1 = "0"
    
    'ENTRADAS en F27:F30
    Range("F27").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("F28").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("F29").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("F30").Select
    ActiveCell.FormulaR1C1 = "0"
    
    'CONTADOR DE INSTRUCCIÓN en C31
    Range("C31").Select
    ActiveCell.FormulaR1C1 = "0"
End Sub

Sub ColorCeldaActivaGrafico()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
End Sub

Sub QuitarColoresGrafico()
    'Limpiar colores de las instrucciones (C9:C29)
    Range("C9:C29").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Limpiar colores de registros y entradas/salidas
    Range("C31,F27:F30,H27:H30,L23,M11,N23,P16").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Restablecer color de fondo de instrucciones
    Range("C9:C29").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    'Restablecer color de fondo de entradas
    Range("F27:F30").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
End Sub

' ========== FUNCIONES DE COMPATIBILIDAD ==========

Sub EjecutarInstruccion()
    ' Redirigir a la nueva función
    EjecutarInstruccionGrafico
End Sub

Sub Resetear()
    ' Redirigir a la nueva función
    ResetearGrafico
End Sub

Sub ColorCeldaActiva()
    ' Redirigir a la nueva función
    ColorCeldaActivaGrafico
End Sub

Sub QuitarColores()
    ' Redirigir a la nueva función
    QuitarColoresGrafico
End Sub

