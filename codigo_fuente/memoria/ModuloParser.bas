Attribute VB_Name = "ModuloParser"
Option Explicit

Public Sub ParsearInstruccion(inst As String)
    Dim partes() As String
    Dim operando1 As String
    Dim operando2 As String
    Dim operando3 As String
    
    partes = Split(inst, " ")
    
    If UBound(partes) >= 0 Then
        Select Case UCase(partes(0))
            Case "MOV"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarMOV operando1, operando2
                End If
            Case "ADD"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarADD operando1, operando2
                End If
            Case "SUB"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarSUB operando1, operando2
                End If
            Case "MUL"
                If UBound(partes) >= 1 Then
                    operando1 = ExtraerOperando(partes(1))
                    EjecutarMUL operando1
                End If
            Case "DIV"
                If UBound(partes) >= 1 Then
                    operando1 = ExtraerOperando(partes(1))
                    EjecutarDIV operando1
                End If
            Case "IMUL"
                If UBound(partes) >= 1 Then
                    operando1 = ExtraerOperando(partes(1))
                    EjecutarIMUL operando1
                End If
            Case "IDIV"
                If UBound(partes) >= 1 Then
                    operando1 = ExtraerOperando(partes(1))
                    EjecutarIDIV operando1
                End If
            Case "AND"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarAND operando1, operando2
                End If
            Case "OR"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarOR operando1, operando2
                End If
            Case "XOR"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarXOR operando1, operando2
                End If
            Case "NOT"
                If UBound(partes) >= 1 Then
                    operando1 = ExtraerOperando(partes(1))
                    EjecutarNOT operando1
                End If
            Case "CMP"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarCMP operando1, operando2
                End If
            Case "TEST"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarTEST operando1, operando2
                End If
            Case "INC"
                If UBound(partes) >= 1 Then
                    operando1 = ExtraerOperando(partes(1))
                    EjecutarINC operando1
                End If
            Case "DEC"
                If UBound(partes) >= 1 Then
                    operando1 = ExtraerOperando(partes(1))
                    EjecutarDEC operando1
                End If
            Case "SHL"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarSHL operando1, operando2
                End If
            Case "SHR"
                If UBound(partes) >= 2 Then
                    operando1 = ExtraerOperando(partes(1))
                    operando2 = ExtraerOperando(partes(2))
                    EjecutarSHR operando1, operando2
                End If
            Case "NOP"
            Case "HLT"
                EIP = MEM_SIZE
            Case Else
                Debug.Print "Instruccion no reconocida: " & partes(0)
        End Select
    End If
End Sub

Private Function ExtraerOperando(texto As String) As String
    Dim temp As String
    temp = Replace(texto, ",", "")
    temp = Trim(temp)
    ExtraerOperando = temp
End Function
' ========== FUNCIONES DE VALIDACIÓN DE OPERANDOS ==========
' Agregar después de la función ExtraerOperando

' Función para verificar si un operando es un registro válido
Private Function EsRegistroValido(registro As String) As Boolean
    Select Case UCase(Trim(registro))
        ' Registros de propósito general de 32 bits
        Case "EAX", "EBX", "ECX", "EDX"
            EsRegistroValido = True
        ' Registros de puntero e índice
        Case "ESP", "EBP", "ESI", "EDI"
            EsRegistroValido = True
        ' Registros de segmento
        Case "CS", "DS", "SS", "ES"
            EsRegistroValido = True
        ' Registro de instrucción
        Case "EIP"
            EsRegistroValido = True
        Case Else
            EsRegistroValido = False
    End Select
End Function

' Función para verificar si un operando es un número válido (decimal o hexadecimal)
Private Function EsNumeroValido(numero As String) As Boolean
    On Error GoTo ErrorHandler
    Dim temp As Long
    Dim numeroLimpio As String
    
    numeroLimpio = UCase(Trim(numero))
    
    ' Verificar si es hexadecimal (formato &H o 0x)
    If Left(numeroLimpio, 2) = "&H" Or Left(numeroLimpio, 2) = "0X" Then
        ' Validar que después del prefijo solo haya dígitos hexadecimales
        Dim hexPart As String
        If Left(numeroLimpio, 2) = "&H" Then
            hexPart = Mid(numeroLimpio, 3)
        Else
            hexPart = Mid(numeroLimpio, 3)
        End If
        
        ' Verificar que solo contenga 0-9, A-F
        Dim i As Integer
        For i = 1 To Len(hexPart)
            Dim char As String
            char = Mid(hexPart, i, 1)
            If Not ((char >= "0" And char <= "9") Or (char >= "A" And char <= "F")) Then
                EsNumeroValido = False
                Exit Function
            End If
        Next i
        
        ' Intentar convertir
        temp = CLng(numeroLimpio)
        EsNumeroValido = True
    Else
        ' Es decimal, verificar que solo contenga dígitos y opcionalmente signo
        If numeroLimpio Like "[-+]#*" Or numeroLimpio Like "#*" Then
            temp = CLng(numeroLimpio)
            EsNumeroValido = True
        Else
            EsNumeroValido = False
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    EsNumeroValido = False
End Function

' Función para validar un operando (puede ser registro, número o dirección de memoria)
Private Function ValidarOperando(operando As String, permitirVacio As Boolean) As Boolean
    Dim operandoLimpio As String
    
    operandoLimpio = Trim(operando)
    
    ' Si está vacío
    If Len(operandoLimpio) = 0 Then
        ValidarOperando = permitirVacio
        Exit Function
    End If
    
    ' Verificar si es un registro válido
    If EsRegistroValido(operandoLimpio) Then
        ValidarOperando = True
        Exit Function
    End If
    
    ' Verificar si es un número válido
    If EsNumeroValido(operandoLimpio) Then
        ValidarOperando = True
        Exit Function
    End If
    
    ' Si llegamos aquí, el operando no es válido
    ValidarOperando = False
End Function

' Función mejorada para validar instrucciones completas
Private Function ValidarInstruccion(inst As String, ByRef mensajeError As String) As Boolean

End Function
    Dim partes() As String
    Dim opcode As String
    Dim operando1 As String
    Dim operando2 As String
    
    mensajeError = ""
    partes = Split(Trim(inst), " ")
    
    ' Verificar que hay al menos un opcode
    If UBound(partes) < 0 Then
        mensajeError = "Instrucción vacía"
        ValidarInstruccionCompleta = False
        Exit Function
    End If
    
    opcode = UCase(partes(0))
    
    ' Validar según el tipo de instrucción
    Select Case opcode
        ' Instrucciones con 2 operandos: destino, origen
        Case "MOV", "ADD", "SUB", "AND", "OR", "XOR", "CMP", "TEST"
            If UBound(partes) < 2 Then
                mensajeError = "Instrucción " & opcode & " requiere 2 operandos"
                ValidarInstruccionCompleta = False
                Exit Function
            End If
            
            operando1 = ExtraerOperando(partes(1))
            operando2 = ExtraerOperando(partes(2))
            
            ' El primer operando debe ser un registro
            If Not EsRegistroValido(operando1) Then
                mensajeError = "Operando destino '" & operando1 & "' no es un registro válido"
                ValidarInstruccionCompleta = False
                Exit Function
            End If
            
            ' El segundo operando puede ser registro o número
            If Not ValidarOperando(operando2, False) Then
                mensajeError = "Operando origen '"


' ========== FUNCIÓN PRINCIPAL CON VALIDACIÓN MEJORADA ==========

Public Sub ParsearYEjecutar(inst As String)
    Dim mensajeError As String
    
    ' Validar la instrucción completa antes de ejecutar
    If ValidarInstruccionCompleta(inst, mensajeError) Then
        ' Si es válida, parsear y ejecutar
        ParsearInstruccion inst
    Else
        ' Si no es válida, mostrar error detallado
        Debug.Print "ERROR: " & mensajeError & " en instrucción: " & inst
        MsgBox "Error de validación:" & vbCrLf & mensajeError & vbCrLf & vbCrLf & _
               "Instrucción: " & inst, vbExclamation, "Error de Sintaxis"
    End If
End Sub
