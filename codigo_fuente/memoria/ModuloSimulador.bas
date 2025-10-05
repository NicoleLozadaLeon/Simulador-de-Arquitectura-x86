Attribute VB_Name = "ModuloSimulador"
' ===== ModuloSimulador =====
Option Explicit

' Variables públicas
Public ALU_Operacion As String
Public ALU_Op1 As String
Public ALU_Op2 As String
Public ALU_Resultado As String
Public DebeDetenerse As Boolean

' Instrucciones del programa
Private ProgramaInstrucciones As Collection
Private InstruccionActual As Integer
Sub InicializarCPU()
    ' Inicializar registros
    EAX = 0: EBX = 0: ECX = 0: EDX = 0
    ESI = 0: EDI = 0: EBP = 0: ESP = STACK_START
    eip = 0
    
    ' Inicializar flags
    ZF = False: SF = False: CF = False: OF = False
    PF = False: AF = False
    
    ' Inicializar ALU
    ALU_Operacion = "-"
    ALU_Op1 = "-"
    ALU_Op2 = "-"
    ALU_Resultado = "-"
    
    DebeDetenerse = False
    InstruccionActual = 0
    Set ProgramaInstrucciones = New Collection
End Sub
Sub CargarProgramaDesdeHoja()
    Call InicializarCPU
    Set ProgramaInstrucciones = New Collection
    
    ' Leer desde hoja Programa
    Dim wsPrograma As Worksheet
    Set wsPrograma = ThisWorkbook.Sheets("Programa")
    
    ' Obtener todo el contenido de A6
    Dim contenido As String
    contenido = Trim(wsPrograma.Range("A6").Value)
    
    ' Si está vacío, mostrar error
    If contenido = "" Then
        MsgBox "No hay programa en la celda A6", vbExclamation
        Exit Sub
    End If
    
    ' Dividir por saltos de línea
    Dim lineas() As String
    lineas = Split(contenido, vbLf) ' Usar vbLf para salto de línea
    
    ' Si no funciona con vbLf, probar con vbCrLf
    If UBound(lineas) = 0 Then
        lineas = Split(contenido, vbCrLf)
    End If
    
    ' Cargar cada línea como instrucción separada
    Dim i As Integer
    For i = 0 To UBound(lineas)
        Dim linea As String
        linea = Trim(lineas(i))
        
        ' Ignorar líneas vacías y comentarios
        If linea <> "" And Left(linea, 1) <> ";" Then
            ProgramaInstrucciones.Add linea
            Debug.Print "Instrucción " & ProgramaInstrucciones.count & ": " & linea
        End If
    Next i
    
    If ProgramaInstrucciones.count > 0 Then
        InstruccionActual = 1
        MsgBox "Programa cargado: " & ProgramaInstrucciones.count & " instrucciones", vbInformation
    Else
        MsgBox "No se encontraron instrucciones válidas", vbExclamation
    End If
    
    ActualizarVistaSimulador
End Sub

Sub EjecutarProgramaCompleto()
    If ProgramaInstrucciones Is Nothing Or ProgramaInstrucciones.count = 0 Then
        CargarProgramaDesdeHoja
    End If
    
    If ProgramaInstrucciones.count = 0 Then Exit Sub
    
    DebeDetenerse = False
    Dim i As Integer
    
    For i = InstruccionActual To ProgramaInstrucciones.count
        If DebeDetenerse Then Exit For
        
        InstruccionActual = i
        EjecutarUnaInstruccion
        
        Dim inicio As Single
        inicio = Timer
        Do While Timer < inicio + 1
            DoEvents
        Loop
    Next i
    
    If Not DebeDetenerse Then
        MsgBox "Ejecución completada", vbInformation
    End If
End Sub
Sub EjecutarPasoAPaso()
    If ProgramaInstrucciones Is Nothing Or ProgramaInstrucciones.count = 0 Then
        CargarProgramaDesdeHoja
    End If
    
    If InstruccionActual > ProgramaInstrucciones.count Then
        MsgBox "Programa terminado", vbInformation
        Exit Sub
    End If
    
    EjecutarUnaInstruccion
    InstruccionActual = InstruccionActual + 1
End Sub
Sub EjecutarUnaInstruccion()
    If ProgramaInstrucciones Is Nothing Or InstruccionActual > ProgramaInstrucciones.count Then
        Exit Sub
    End If
    
    Dim instruccion As String
    instruccion = ProgramaInstrucciones(InstruccionActual)
    
    ' Mostrar en consola para debug
    Debug.Print "Ejecutando instrucción " & InstruccionActual & ": " & instruccion
    
    ' Configurar display ALU
    ConfigurarDisplayALU instruccion
    
    ' Ejecutar instrucción
    ModuloParser.ParsearYEjecutar instruccion
    
    ' Actualizar vista
    ActualizarVistaSimulador
End Sub
Sub ConfigurarDisplayALU(instruccion As String)
    Dim partes() As String
    partes = Split(Trim(instruccion), " ", 3) ' Dividir máximo en 3 partes
    
    ALU_Operacion = "-"
    ALU_Op1 = "-"
    ALU_Op2 = "-"
    ALU_Resultado = "-"
    
    If UBound(partes) >= 0 Then
        ALU_Operacion = UCase(partes(0))
    End If
    
    If UBound(partes) >= 1 Then
        ALU_Op1 = ExtraerOperando(partes(1))
    End If
    
    If UBound(partes) >= 2 Then
        ALU_Op2 = ExtraerOperando(partes(2))
    End If
    
    Select Case ALU_Operacion
        Case "MOV"
            If ALU_Op1 <> "-" Then
                ALU_Resultado = ALU_Op1 & " = " & ALU_Op2
            End If
        Case "ADD"
            If ALU_Op1 <> "-" Then
                ALU_Resultado = ALU_Op1 & " + " & ALU_Op2
            End If
        Case "SUB"
            If ALU_Op1 <> "-" Then
                ALU_Resultado = ALU_Op1 & " - " & ALU_Op2
            End If
        Case "MUL", "IMUL"
            ALU_Resultado = "EAX * " & ALU_Op1
        Case "DIV", "IDIV"
            ALU_Resultado = "EAX / " & ALU_Op1
        Case Else
            ALU_Resultado = ALU_Operacion
    End Select
End Sub
Function ExtraerOperando(texto As String) As String
    If texto = "" Then
        ExtraerOperando = ""
        Exit Function
    End If
    
    ' Eliminar comas y espacios extra
    texto = Trim(Replace(texto, ",", ""))
    
    ' Si termina con MOV (error de parsing), quitarlo
    If UCase(Right(texto, 3)) = "MOV" Then
        texto = Trim(Left(texto, Len(texto) - 3))
    End If
    
    ExtraerOperando = UCase(texto)
End Function
Sub ReiniciarTodo()
    DebeDetenerse = True
    Call InicializarCPU
    ActualizarVistaSimulador
    MsgBox "Sistema reiniciado", vbInformation
End Sub
Sub ActualizarVistaSimulador()
    On Error Resume Next
    Dim wsSimulador As Worksheet
    Set wsSimulador = ThisWorkbook.Sheets("Simulador")
    
    With wsSimulador
        ' Actualizar registros
        .Range("B5").Value = EAX
        .Range("B6").Value = EBX
        .Range("B7").Value = ECX
        .Range("B8").Value = EDX
        .Range("B9").Value = ESI
        .Range("B10").Value = EDI
        .Range("B11").Value = EBP
        .Range("B12").Value = ESP
        
        ' Actualizar flags
        .Range("E5").Value = IIf(ZF, "1", "0")
        .Range("E6").Value = IIf(SF, "1", "0")
        .Range("E7").Value = IIf(CF, "1", "0")
        .Range("E8").Value = IIf(OF, "1", "0")
        .Range("E9").Value = IIf(PF, "1", "0")
        .Range("E10").Value = IIf(AF, "1", "0")
        
        ' Actualizar ALU
        .Range("H5").Value = ALU_Operacion
        .Range("H6").Value = ALU_Op1
        .Range("H7").Value = ALU_Op2
        .Range("H8").Value = ALU_Resultado
        
        ' Actualizar Unidad de Control - CORREGIDO
        If ProgramaInstrucciones Is Nothing Or ProgramaInstrucciones.count = 0 Then
            .Range("L5").Value = "Idle"
            .Range("L6").Value = "-"
        ElseIf InstruccionActual > ProgramaInstrucciones.count Then
            .Range("L5").Value = "Complete"
            .Range("L6").Value = "Programa terminado"
        Else
            .Range("L5").Value = "Execute"
            ' Mostrar SOLO la instrucción actual, no todo el programa
            .Range("L6").Value = ProgramaInstrucciones(InstruccionActual)
        End If
    End With
End Sub
Private Function EsNumero(texto As String) As Boolean
    On Error GoTo ErrorHandler
    Dim temp As String
    temp = UCase(Trim(texto))
    
    ' Si es hexadecimal
    If Left(temp, 2) = "0X" Then
        temp = "&H" & Mid(temp, 3)
    End If
    
    ' Verificar si es número
    If IsNumeric(temp) Then
        EsNumero = True
    Else
        EsNumero = False
    End If
    Exit Function
    
ErrorHandler:
    EsNumero = False
End Function
