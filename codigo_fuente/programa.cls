VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "programa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ===== HOJA: Programa =====
Option Explicit

Private Sub Worksheet_Activate()
    Call ConfigurarHojaPrograma
End Sub

Public Sub ConfigurarHojaPrograma()
    With Me
        .Cells.Clear
        
        ' Título
        .Range("A1:J1").Merge
        .Range("A1").value = "INGRESA TU PROGRAMA EN ENSAMBLADOR x86"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(60, 100, 160)
        .Range("A1").Font.Color = RGB(255, 255, 255)
        
        ' Instrucciones de ejemplo
        .Range("A3").value = "Ejemplo (usa Alt+Enter para saltos de línea):"
        .Range("A4").value = "mov eax, 10" & Chr(10) & "mov ebx, 3" & Chr(10) & "add eax, ebx"
        .Range("A4").Font.Italic = True
        .Range("A4").Font.Size = 9
        .Range("A4").Font.Color = RGB(80, 80, 80)
        .Range("A4").WrapText = True
        
        ' Cuadro de código (EN BLANCO)
        .Range("A6:J20").Merge
        .Range("A6").Borders.LineStyle = xlContinuous
        .Range("A6").Interior.Color = RGB(250, 250, 255)
        .Range("A6").Font.Name = "Courier New"
        .Range("A6").Font.Size = 10
        .Range("A6").WrapText = True
        .Range("A6").VerticalAlignment = xlTop
        .Range("A6").value = "" ' EN BLANCO
        
        ' Botones
        Dim btnEjecutar As Button, btnReiniciar As Button, btnCache As Button, btnRAM As Button
        On Error Resume Next
        .Buttons.Delete
        On Error GoTo 0
        
        Set btnEjecutar = .Buttons.Add(100, 350, 120, 30)
        With btnEjecutar
            .Name = "btnEjecutar"
            .Caption = "?? Ejecutar"
            .OnAction = "'Programa.Programa_Ejecutar'"
            .Font.Bold = True
        End With
        
        Set btnReiniciar = .Buttons.Add(250, 350, 120, 30)
        With btnReiniciar
            .Name = "btnReiniciar"
            .Caption = "?? Reiniciar"
            .OnAction = "'Programa.Programa_Reiniciar'"
            .Font.Bold = True
        End With
        
        ' Botón para Simulador de Caché
        Set btnCache = .Buttons.Add(400, 350, 150, 30)
        With btnCache
            .Name = "btnCache"
            .Caption = "?? Simular Caché"
            .OnAction = "'Programa.Programa_EjecutarCache'"
            .Font.Bold = True
        End With
        
        ' NUEVO: Botón para Simulador de RAM
        Set btnRAM = .Buttons.Add(570, 350, 150, 30)
        With btnRAM
            .Name = "btnRAM"
            .Caption = "?? Simular RAM"
            .OnAction = "'Programa.Programa_EjecutarRAM'"
            .Font.Bold = True
        End With
        
        .Columns("A:J").ColumnWidth = 12
        .ROWS(6).RowHeight = 120
    End With
End Sub
' En la hoja Programa, ACTUALIZA la subrutina Programa_Ejecutar:

Sub Programa_Ejecutar()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' 1. Transferir código a la hoja ProgramaNASM para el simulador de caché
    TransferirProgramaACache
    
    ' 2. Transferir código a la hoja CodigoVM para el simulador de memoria virtual
    TransferirProgramaAVirtualMemory
    
    ' 3. Transferir código para el pipeline con riesgos
    TransferirProgramaAPipelineRiesgos
    
    ' 4. EJECUTAR SIMULADOR DE RAM (NUEVO)
    RAM_Simulator.IniciarSimuladorRAM
    
    ' 5. Ejecutar Pipeline PRINCIPAL (el que ya funciona)
    Pipeline.IniciarPipeline
    Pipeline.EjecutarPipelineCompleto
    
    ' 6. Ejecutar Pipeline con DETECCIÓN DE RIESGOS (nuevo módulo)
    PipelineRiesgos.IniciarSimuladorConRiesgos
    PipelineRiesgos.EjecutarPipelineRiesgosCompleto
    
    ' 7. Ejecutar Simulador tradicional
    ModuloSimulador.CargarProgramaDesdeHoja
    Sheets("Simulador").Activate
    ModuloSimulador.EjecutarProgramaCompleto
    
    ' 8. Ejecutar Simulador de Memoria Virtual
    Virtual_Memory_Simulator.InitializeVirtualMemorySimulator
    Virtual_Memory_Simulator.ExecuteFullProgramVM
    
    ' 9. Ejecutar Simulador de Caché existente
    Cache_Simulator.InitializeCacheSimulator
    Cache_Simulator.ExecuteFullProgramCache
    
    ' 10. EJECUTAR NUEVO SIMULADOR LRU
    EjecutarSimuladorLRU
    
    ' 11. EJECUTAR PROGRAMA EN RAM (NUEVO)
    RAM_Simulator.EjecutarProgramaRAMCompleto
    
    ' 12. Activar hoja de pipeline con riesgos para mostrar resultados
    Sheets("PipelineRiesgos").Activate
    
    Application.ScreenUpdating = True
    
    MsgBox "Ejecución completada en todos los simuladores." & vbCrLf & _
           "Incluyendo nuevo simulador de RAM y política de reemplazo LRU.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error al ejecutar: " & Err.Description, vbCritical
End Sub

' NUEVO: Ejecutar programa con simulador de caché
Sub Programa_EjecutarCache()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Primero, transferir el código a la hoja ProgramaNASM
    TransferirProgramaACache
    
    ' 1. Ejecutar Simulador de Caché existente
    Cache_Simulator.InitializeCacheSimulator
    Cache_Simulator.ExecuteFullProgramCache
    
    ' 2. Ejecutar Simulador LRU para comparar políticas
    LRU_Simulador.InitializeCache 2, 16, 2
    GenerarAccesosLRUDesdePrograma
    
    ' Crear hoja de comparación entre políticas
    CrearComparacionPoliticasCache
    
    ' Activar la hoja de comparación
    Sheets("ComparacionCache").Activate
    
    Application.ScreenUpdating = True
    
    MsgBox "Simulación de caché completada." & vbCrLf & _
           "Se ejecutaron ambos simuladores para comparar políticas." & vbCrLf & _
           "Revise las hojas 'Cache', 'EstadoCaché' y 'ComparacionCache'.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error al ejecutar simulador de caché: " & Err.Description & vbCrLf & _
           "Asegúrese de que todas las hojas necesarias existan.", vbCritical
End Sub

' FUNCIÓN CORREGIDA: Crear comparación entre políticas de caché
Sub CrearComparacionPoliticasCache()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ComparacionCache")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "ComparacionCache"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    ' Título
    With ws.Range("A1:E1")
        .Merge
        .value = "?? COMPARACIÓN DE POLÍTICAS DE CACHÉ - LRU vs OTRAS"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(70, 130, 180)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Descripción
    ws.Range("A3").value = "Esta comparación muestra las diferencias entre la política LRU (Least Recently Used) y otras políticas de reemplazo."
    ws.Range("A3").Font.Italic = True
    ws.Range("A3").WrapText = True
    
    ' Características de LRU
    ws.Range("A5").value = "?? POLÍTICA LRU (LEAST RECENTLY USED)"
    ws.Range("A5").Font.Bold = True
    ws.Range("A5").Font.Size = 12
    
    ws.Range("A6").value = "Ventajas:"
    ws.Range("A6").Font.Bold = True
    ws.Range("A7").value = "• Reemplaza el bloque que no se ha usado por más tiempo"
    ws.Range("A8").value = "• Buen rendimiento para patrones de acceso con localidad temporal"
    ws.Range("A9").value = "• Evita reemplazar bloques que se usan frecuentemente"
    
    ws.Range("A11").value = "Desventajas:"
    ws.Range("A11").Font.Bold = True
    ws.Range("A12").value = "• Requiere mantener contadores de uso para cada bloque"
    ws.Range("A13").value = "• Mayor complejidad de implementación"
    ws.Range("A14").value = "• Puede no ser óptimo para algunos patrones de acceso"
    
    ' Otras políticas
    ws.Range("C5").value = "?? OTRAS POLÍTICAS DE REEMPLAZO"
    ws.Range("C5").Font.Bold = True
    ws.Range("C5").Font.Size = 12
    
    ws.Range("C6").value = "FIFO (First-In, First-Out):"
    ws.Range("C6").Font.Bold = True
    ws.Range("C7").value = "• Reemplaza el bloque más antiguo"
    ws.Range("C8").value = "• Simple de implementar"
    ws.Range("C9").value = "• Puede reemplazar bloques usados frecuentemente"
    
    ws.Range("C11").value = "LFU (Least Frequently Used):"
    ws.Range("C11").Font.Bold = True
    ws.Range("C12").value = "• Reemplaza el bloque menos frecuentemente usado"
    ws.Range("C13").value = "• Bueno para patrones de acceso estables"
    ws.Range("C14").value = "• Puede mantener bloques 'viejos' que ya no se usan"
    
    ws.Range("C16").value = "Random:"
    ws.Range("C16").Font.Bold = True
    ws.Range("C17").value = "• Reemplaza un bloque aleatorio"
    ws.Range("C18").value = "• Muy simple de implementar"
    ws.Range("C19").value = "• Rendimiento impredecible"
    
    ' Estadísticas comparativas - CORREGIDO: usar funciones públicas
    ws.Range("A20").value = "?? ESTADÍSTICAS COMPARATIVAS"
    ws.Range("A20").Font.Bold = True
    ws.Range("A20").Font.Size = 12
    
    ws.Range("A21").value = "Política"
    ws.Range("B21").value = "Tasa de Hit"
    ws.Range("C21").value = "Tasa de Miss"
    ws.Range("D21").value = "Reemplazos"
    ws.Range("E21").value = "Eficiencia"
    
    ws.Range("A22").value = "LRU"
    ws.Range("B22").value = Format(LRU_Simulador.GetHitRate(), "0.00%")
    ws.Range("C22").value = Format(LRU_Simulador.GetMissRate(), "0.00%")
    ws.Range("D22").value = LRU_Simulador.GetReplacementCount()
    ws.Range("E22").value = "Alta para localidad temporal"
    
    ' Formato de la tabla
    Dim headerRange As Range
    Set headerRange = ws.Range("A21:E21")
    headerRange.Font.Bold = True
    headerRange.Interior.Color = RGB(200, 200, 200)
    
    ws.Range("A21:E22").Borders.LineStyle = xlContinuous
    
    ' Consejos para análisis
    ws.Range("A25").value = "?? CONSEJOS PARA EL ANÁLISIS"
    ws.Range("A25").Font.Bold = True
    ws.Range("A25").Font.Size = 12
    
    ws.Range("A26").value = "1. Revise la hoja 'EstadoCaché' para ver el estado actual de la caché LRU"
    ws.Range("A27").value = "2. Consulte 'LogCaché' para el historial detallado de accesos"
    ws.Range("A28").value = "3. Verifique 'Estadísticas' para métricas de rendimiento"
    ws.Range("A29").value = "4. Compare con el simulador de caché tradicional en la hoja 'Cache'"
    
    ' Agregar información adicional sobre la simulación
    ws.Range("A31").value = "?? INFORMACIÓN DE LA SIMULACIÓN LRU"
    ws.Range("A31").Font.Bold = True
    ws.Range("A31").Font.Size = 12
    
    ws.Range("A32").value = "Total de accesos: " & LRU_Simulador.GetAccessCounter()
    ws.Range("A33").value = "Hits: " & LRU_Simulador.GetHitCount()
    ws.Range("A34").value = "Misses: " & LRU_Simulador.GetMissCount()
    ws.Range("A35").value = "Reemplazos LRU: " & LRU_Simulador.GetReplacementCount()
    
    ws.Columns("A:A").ColumnWidth = 35
    ws.Columns("B:E").ColumnWidth = 15
    ws.Columns("C:C").ColumnWidth = 25
    
    ' Autoajustar filas
    ws.ROWS.AutoFit
End Sub

Sub Programa_Reiniciar()
    On Error GoTo ErrorHandler
    ' Reiniciar todos los sistemas
    ModuloSimulador.ReiniciarTodo
    Pipeline.ReiniciarPipeline
    
    ' También reiniciar los simuladores de caché y RAM
    On Error Resume Next
    Cache_Simulator.ResetCacheSimulator
    Virtual_Memory_Simulator.ResetVirtualMemory
    LRU_Simulador.ResetCacheSimulator
    RAM_Simulator.ReiniciarSimuladorRAM ' NUEVO: Reiniciar RAM
    On Error GoTo 0
    
    Sheets("Simulador").Activate
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al reiniciar: " & Err.Description, vbCritical
End Sub

' FUNCIÓN PARA TRANSFERIR EL PROGRAMA A LA HOJA CACHE
Private Sub TransferirProgramaACache()
    Dim wsPrograma As Worksheet, wsNASM As Worksheet
    Dim codigo As String
    Dim lineas() As String
    Dim i As Integer
    
    Set wsPrograma = Sheets("Programa")
    
    ' Crear o acceder a la hoja ProgramaNASM
    On Error Resume Next
    Set wsNASM = Sheets("ProgramaNASM")
    If wsNASM Is Nothing Then
        Set wsNASM = Worksheets.Add
        wsNASM.Name = "ProgramaNASM"
    End If
    On Error GoTo 0
    
    ' Obtener el código del cuadro de texto
    codigo = wsPrograma.Range("A6").value
    If codigo = "" Then
        MsgBox "No hay código para ejecutar. Por favor, ingresa un programa en ensamblador.", vbExclamation
        Exit Sub
    End If
    
    ' Dividir en líneas
    lineas = Split(codigo, Chr(10))
    
    ' Limpiar hoja NASM y escribir el código
    wsNASM.Cells.Clear
    wsNASM.Range("A1").value = "Código NASM (Generado desde Programa)"
    wsNASM.Range("A1").Font.Bold = True
    
    ' Escribir cada línea - CONVERTIR A FORMATO NASM COMPATIBLE CON RAM
    Dim lineCount As Integer
    lineCount = 2
    
    ' Sección de datos
    wsNASM.Cells(lineCount, 1).value = "section .data"
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = "    num1 dd 10"
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = "    num2 dd 20"
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = "    result dd 0"
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = ""
    lineCount = lineCount + 1
    
    ' Sección de código
    wsNASM.Cells(lineCount, 1).value = "section .text"
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = "    global _start"
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = ""
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = "_start:"
    lineCount = lineCount + 1
    
    ' Convertir instrucciones simples a formato NASM compatible
    For i = 0 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            Dim instruction As String
            instruction = ConvertToNASMFormatRAM(Trim(lineas(i))) ' Nueva función mejorada
            wsNASM.Cells(lineCount, 1).value = "    " & instruction
            lineCount = lineCount + 1
        End If
    Next i
    
    ' Agregar instrucciones de terminación
    wsNASM.Cells(lineCount, 1).value = "    mov eax, 1"
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = "    xor ebx, ebx"
    lineCount = lineCount + 1
    wsNASM.Cells(lineCount, 1).value = "    int 0x80"
    
    wsNASM.Columns("A").AutoFit
End Sub

Private Function ConvertToNASMFormatRAM(instruction As String) As String
    Dim parts() As String
    parts = Split(instruction, " ")
    
    If UBound(parts) < 1 Then
        ConvertToNASMFormatRAM = instruction
        Exit Function
    End If
    
    Dim Opcode As String
    Opcode = LCase(parts(0))
    
    ' Convertir instrucciones a formato NASM que el simulador de RAM pueda parsear
    Select Case Opcode
        Case "mov"
            If UBound(parts) >= 2 Then
                If IsNumeric(Replace(parts(2), ",", "")) Then
                    ' mov eax, 10 -> mov eax, [num1]
                    ConvertToNASMFormatRAM = "mov " & parts(1) & ", [num1]"
                Else
                    ConvertToNASMFormatRAM = instruction
                End If
            Else
                ConvertToNASMFormatRAM = instruction
            End If
            
        Case "add", "sub", "mul", "div"
            If UBound(parts) >= 2 Then
                If IsNumeric(Replace(parts(2), ",", "")) Then
                    ' add eax, 5 -> add eax, [num2]
                    ConvertToNASMFormatRAM = Opcode & " " & parts(1) & ", [num2]"
                Else
                    ConvertToNASMFormatRAM = instruction
                End If
            Else
                ConvertToNASMFormatRAM = instruction
            End If
            
        Case Else
            ConvertToNASMFormatRAM = instruction
    End Select
End Function
' NUEVA FUNCIÓN: Transferir programa a memoria virtual
Private Sub TransferirProgramaAVirtualMemory()
    Dim wsPrograma As Worksheet, wsVM As Worksheet
    Dim codigo As String
    Dim lineas() As String
    Dim i As Integer
    
    Set wsPrograma = Sheets("Programa")
    
    ' Crear o acceder a la hoja CodigoVM
    On Error Resume Next
    Set wsVM = Sheets("CodigoVM")
    If wsVM Is Nothing Then
        Set wsVM = Worksheets.Add
        wsVM.Name = "CodigoVM"
    End If
    On Error GoTo 0
    
    ' Obtener el código del cuadro de texto
    codigo = wsPrograma.Range("A6").value
    If codigo = "" Then
        MsgBox "No hay código para ejecutar. Por favor, ingresa un programa en ensamblador.", vbExclamation
        Exit Sub
    End If
    
    ' Dividir en líneas
    lineas = Split(codigo, Chr(10))
    
    ' Limpiar hoja VM y escribir el código
    wsVM.Cells.Clear
    wsVM.Range("A1").value = "Programa Ensamblador (Generado desde Programa)"
    wsVM.Range("A1").Font.Bold = True
    
    ' Escribir cada línea en formato compatible con memoria virtual
    Dim lineCount As Integer
    lineCount = 2
    
    For i = 0 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            Dim instruction As String
            instruction = ConvertToVMFormat(Trim(lineas(i)))
            If instruction <> "" Then
                wsVM.Cells(lineCount, 1).value = instruction
                lineCount = lineCount + 1
            End If
        End If
    Next i
    
    ' Agregar instrucción de terminación si no existe
    If lineCount > 2 Then
        wsVM.Cells(lineCount, 1).value = "HLT"
    End If
    
    wsVM.Columns("A").AutoFit
End Sub

' FUNCIÓN PARA CONVERTIR INSTRUCCIONES SIMPLES A FORMATO NASM
Private Function ConvertToNASMFormat(instruction As String) As String
    Dim parts() As String
    parts = Split(instruction, " ")
    
    If UBound(parts) < 1 Then
        ConvertToNASMFormat = instruction
        Exit Function
    End If
    
    Dim Opcode As String
    Opcode = LCase(parts(0))
    
    ' Convertir instrucciones con constantes a accesos a memoria para demostrar caché
    Select Case Opcode
        Case "mov"
            If UBound(parts) >= 2 Then
                If IsNumeric(Replace(parts(2), ",", "")) Then
                    ' mov eax, 10 -> mov eax, [num1] (para generar accesos a memoria)
                    ConvertToNASMFormat = "mov " & parts(1) & ", [num1]"
                Else
                    ConvertToNASMFormat = instruction
                End If
            Else
                ConvertToNASMFormat = instruction
            End If
            
        Case "add"
            If UBound(parts) >= 2 Then
                If IsNumeric(Replace(parts(2), ",", "")) Then
                    ' add eax, 5 -> add eax, [num2]
                    ConvertToNASMFormat = "add " & parts(1) & ", [num2]"
                Else
                    ConvertToNASMFormat = instruction
                End If
            Else
                ConvertToNASMFormat = instruction
            End If
            
        Case Else
            ConvertToNASMFormat = instruction
    End Select
End Function

' NUEVA FUNCIÓN: Convertir instrucciones a formato compatible con memoria virtual
Private Function ConvertToVMFormat(instruction As String) As String
    Dim parts() As String
    parts = Split(instruction, " ")
    
    If UBound(parts) < 1 Then
        ConvertToVMFormat = instruction
        Exit Function
    End If
    
    Dim Opcode As String
    Opcode = UCase(parts(0))
    
    ' Convertir sintaxis x86 a formato simple para memoria virtual
    Select Case Opcode
        Case "MOV"
            If UBound(parts) >= 2 Then
                ' Convertir "mov eax, 10" a "MOV AX 10"
                Dim reg As String, value As String
                reg = ConvertRegisterVM(parts(1))
                value = Replace(parts(2), ",", "")
                ConvertToVMFormat = "MOV " & reg & " " & value
            Else
                ConvertToVMFormat = instruction
            End If
            
        Case "ADD"
            If UBound(parts) >= 2 Then
                Dim reg1 As String, reg2 As String
                reg1 = ConvertRegisterVM(parts(1))
                reg2 = ConvertRegisterVM(Replace(parts(2), ",", ""))
                ConvertToVMFormat = "ADD " & reg1 & " " & reg2
            Else
                ConvertToVMFormat = instruction
            End If
            
        Case "SUB"
            If UBound(parts) >= 2 Then
                reg1 = ConvertRegisterVM(parts(1))
                reg2 = ConvertRegisterVM(Replace(parts(2), ",", ""))
                ConvertToVMFormat = "SUB " & reg1 & " " & reg2
            Else
                ConvertToVMFormat = instruction
            End If
            
        Case Else
            ConvertToVMFormat = instruction
    End Select
End Function

' FUNCIÓN AUXILIAR: Convertir registros x86 a formato VM
Private Function ConvertRegisterVM(reg As String) As String
    Dim cleanReg As String
    cleanReg = UCase(Replace(reg, ",", ""))
    
    Select Case cleanReg
        Case "EAX", "AX": ConvertRegisterVM = "AX"
        Case "EBX", "BX": ConvertRegisterVM = "BX"
        Case "ECX", "CX": ConvertRegisterVM = "CX"
        Case "EDX", "DX": ConvertRegisterVM = "DX"
        Case "ESI", "SI": ConvertRegisterVM = "SI"
        Case "EDI", "DI": ConvertRegisterVM = "DI"
        Case Else: ConvertRegisterVM = cleanReg
    End Select
End Function

' FUNCIÓN AUXILIAR PARA VERIFICAR SI ES NÚMERO
Private Function IsNumeric(ByVal value As String) As Boolean
    On Error GoTo ErrorHandler
    Dim d As Double
    d = CDbl(value)
    IsNumeric = True
    Exit Function
ErrorHandler:
    IsNumeric = False
End Function

Private Sub TransferirProgramaAPipelineRiesgos()
    Dim wsPrograma As Worksheet, wsPipeline As Worksheet
    Dim codigo As String
    Dim lineas() As String
    Dim i As Integer
    
    Set wsPrograma = Sheets("Programa")
    
    ' Crear o acceder a la hoja de pipeline
    On Error Resume Next
    Set wsPipeline = Sheets("PipelineCodigo")
    If wsPipeline Is Nothing Then
        Set wsPipeline = Worksheets.Add
        wsPipeline.Name = "PipelineCodigo"
    End If
    On Error GoTo 0
    
    ' Obtener el código del cuadro de texto
    codigo = wsPrograma.Range("A6").value
    If codigo = "" Then
        ' Usar programa de ejemplo que genera riesgos REALES
        ' RAW hazard: R1 no está listo en SUB R4, R1, R5
        ' RAW hazard: R1 no está listo en ADD R9, R1, R10
        ' RAW hazard: R4 y R1 en SUB R13, R4, R1
        codigo = "ADD R1, R2, R3" & Chr(10) & _
                 "SUB R4, R1, R5" & Chr(10) & _
                 "MUL R6, R7, R8" & Chr(10) & _
                 "ADD R9, R1, R10" & Chr(10) & _
                 "MOV R11, R12" & Chr(10) & _
                 "SUB R13, R4, R1"
    End If
    
    ' Dividir en líneas
    lineas = Split(codigo, Chr(10))
    
    ' Limpiar hoja y escribir el código
    wsPipeline.Cells.Clear
    wsPipeline.Range("A1").value = "Código para Pipeline con Riesgos"
    wsPipeline.Range("A1").Font.Bold = True
    
    ' Escribir cada línea convertida a formato pipeline
    Dim lineCount As Integer
    lineCount = 2
    
    For i = 0 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            Dim instruction As String
            instruction = ConvertToPipelineFormat(Trim(lineas(i)))
            If instruction <> "" Then
                wsPipeline.Cells(lineCount, 1).value = instruction
                lineCount = lineCount + 1
            End If
        End If
    Next i
    
    wsPipeline.Columns("A").AutoFit
End Sub
Private Function ConvertToPipelineFormat(instruction As String) As String
    Dim parts() As String
    parts = Split(instruction, " ")
    
    If UBound(parts) < 1 Then
        ConvertToPipelineFormat = instruction
        Exit Function
    End If
    
    Dim Opcode As String
    Opcode = UCase(parts(0))
    
    ' Convertir sintaxis a formato pipeline
    Select Case Opcode
        Case "MOV"
            If UBound(parts) >= 2 Then
                Dim dest As String, src As String
                dest = ConvertRegisterPipeline(parts(1))
                src = Replace(parts(2), ",", "")
                If IsNumeric(src) Then
                    ConvertToPipelineFormat = "MOV " & dest & ", " & src
                Else
                    ConvertToPipelineFormat = "MOV " & dest & ", " & ConvertRegisterPipeline(src)
                End If
            End If
            
        Case "ADD", "SUB", "MUL", "DIV"
            If UBound(parts) >= 3 Then
                dest = ConvertRegisterPipeline(parts(1))
                Dim src1 As String, src2 As String
                src1 = ConvertRegisterPipeline(Replace(parts(2), ",", ""))
                src2 = ConvertRegisterPipeline(parts(3))
                ConvertToPipelineFormat = Opcode & " " & dest & ", " & src1 & ", " & src2
            End If
            
        Case Else
            ConvertToPipelineFormat = instruction
    End Select
End Function

' FUNCIÓN AUXILIAR: Convertir registros a formato pipeline
Private Function ConvertRegisterPipeline(reg As String) As String
    Dim cleanReg As String
    cleanReg = UCase(Replace(reg, ",", ""))
    
    Select Case cleanReg
        Case "EAX", "AX", "RAX": ConvertRegisterPipeline = "R1"
        Case "EBX", "BX", "RBX": ConvertRegisterPipeline = "R2"
        Case "ECX", "CX", "RCX": ConvertRegisterPipeline = "R3"
        Case "EDX", "DX", "RDX": ConvertRegisterPipeline = "R4"
        Case "ESI", "SI", "RSI": ConvertRegisterPipeline = "R5"
        Case "EDI", "DI", "RDI": ConvertRegisterPipeline = "R6"
        Case Else
            If left(cleanReg, 1) = "R" Then
                ConvertRegisterPipeline = cleanReg
            Else
                ConvertRegisterPipeline = cleanReg
            End If
    End Select
End Function


' NUEVA FUNCIÓN: Ejecutar simulador LRU
Sub EjecutarSimuladorLRU()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Inicializar el simulador LRU con configuración por defecto
    LRU_Simulador.InitializeCache 2, 16, 2 ' 2KB, bloques de 16 bytes, 2-way
    
    ' Generar secuencia de accesos de prueba basada en el programa
    GenerarAccesosLRUDesdePrograma
    
    ' Mostrar resultados
    On Error Resume Next
    Sheets("EstadoCaché").Activate
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    MsgBox "Simulación de política LRU completada." & vbCrLf & _
           "Revise las hojas 'EstadoCaché', 'LogCaché' y 'Estadísticas' para ver los resultados.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error al ejecutar simulador LRU: " & Err.Description, vbCritical
End Sub
' FUNCIÓN PARA GENERAR ACCESOS LRU DESDE EL PROGRAMA
' FUNCIÓN MEJORADA PARA GENERAR ACCESOS LRU DESDE EL PROGRAMA
' FUNCIÓN MEJORADA PARA GENERAR ACCESOS LRU DESDE EL PROGRAMA
' FUNCIÓN MEJORADA PARA GENERAR ACCESOS LRU DESDE EL PROGRAMA
' FUNCIÓN CORREGIDA PARA GENERAR ACCESOS LRU DESDE EL PROGRAMA
Sub GenerarAccesosLRUDesdePrograma()
    Dim wsPrograma As Worksheet
    Dim codigo As String
    Dim lineas() As String
    Dim i As Long
    Dim addressBase As Long
    Dim accessCount As Long
    
    Set wsPrograma = Sheets("Programa")
    codigo = wsPrograma.Range("A6").value
    
    ' Si no hay código, usar secuencia de prueba que DEMUESTRE LRU
    If Trim(codigo) = "" Then
        ' Secuencia de prueba diseñada específicamente para demostrar LRU
        Dim testAddresses As Variant
        testAddresses = Array(0, 16, 32, 48, 0, 64, 16, 80, 32, 96, 0, 112, 16, 128, 32, 144)
        
        For i = LBound(testAddresses) To UBound(testAddresses)
            ' Convertir explícitamente a Long
            LRU_Simulador.AccessMemory CLng(testAddresses(i))
        Next i
    Else
        ' Generar accesos basados en el código del programa
        lineas = Split(codigo, Chr(10))
        addressBase = 0
        accessCount = 0
        
        For i = 0 To UBound(lineas)
            If Trim(lineas(i)) <> "" Then
                ' Generar múltiples accesos por instrucción para simular mejor
                LRU_Simulador.AccessMemory addressBase + accessCount * 16    ' Acceso a instrucción
                LRU_Simulador.AccessMemory addressBase + accessCount * 16 + 4 ' Acceso a datos
                LRU_Simulador.AccessMemory addressBase + accessCount * 16 + 8 ' Otro acceso
                accessCount = accessCount + 1
            End If
        Next i
        
        ' Si hay pocas instrucciones, agregar más accesos para demostrar LRU
        If accessCount < 6 Then
            For i = accessCount To 12
                LRU_Simulador.AccessMemory i * 16
                LRU_Simulador.AccessMemory i * 16 + 32 ' Patrón que fuerza reemplazos
            Next i
        End If
    End If
    
    ' Log final - CORREGIDO: usar la función pública
    LRU_Simulador.LogMessage "? Simulación LRU completada. Total accesos: " & LRU_Simulador.GetAccessCounter()
End Sub

' NUEVO: Ejecutar programa con simulador de RAM
Sub Programa_EjecutarRAM()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Primero, transferir el código a la hoja ProgramaNASM
    TransferirProgramaACache
    
    ' Inicializar y ejecutar el simulador de RAM
    RAM_Simulator.IniciarSimuladorRAM
    RAM_Simulator.EjecutarProgramaRAMCompleto
    
    ' Activar la hoja de RAM para ver los resultados
    Sheets("RAM").Activate
    
    Application.ScreenUpdating = True
    
    MsgBox "Simulación de RAM completada." & vbCrLf & _
           "Revise la hoja 'RAM' para ver cómo se carga el programa en memoria.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error al ejecutar simulador de RAM: " & Err.Description & vbCrLf & _
           "Asegúrese de que todas las hojas necesarias existan.", vbCritical
End Sub
