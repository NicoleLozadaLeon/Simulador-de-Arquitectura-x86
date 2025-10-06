Attribute VB_Name = "Cache_Simulator"
' Módulo: Cache_Simulator
' Descripción: Simulador de memoria caché con visualización completa

Option Explicit

' Constantes para la configuración de la caché
Private Const CACHE_SIZE As Integer = 16      ' 16 líneas de caché
Private Const CACHE_LINE_SIZE As Integer = 4  ' 4 bytes por línea
Private Const RAM_SIZE As Integer = 256
Private Const ROWS As Integer = 16
Private Const COLS As Integer = 16

' Constantes para la caché 2-way
Private Const CACHE_2WAY_SETS As Integer = 32     ' 32 conjuntos
Private Const CACHE_2WAY_BLOCKS_PER_SET As Integer = 2  ' 2 vías
Private Const CACHE_2WAY_BLOCK_SIZE As Integer = 16     ' 16 bytes por bloque
Private Const CACHE_2WAY_TOTAL_SIZE As Integer = 1024   ' 1 KB total

' Estructura de línea de caché para mapeo directo
Private Type cacheLine
    Valid As Boolean
    tag As Long
    Data(0 To CACHE_LINE_SIZE - 1) As Byte
    address As Long
    accessCount As Integer
End Type

' Estructura para caché 2-way
Private Type cacheBlock2Way
    Valid As Boolean
    tag As Long
    Data(0 To CACHE_2WAY_BLOCK_SIZE - 1) As Byte
    address As Long
    LastUsed As Long
    accessCount As Integer
End Type

' Tipo para instrucciones
Private Type AssemblyInstruction
    address As Long
    OriginalLine As String
    Opcode As String
    Operand1 As String
    Operand2 As String
    bytes As String
    Length As Integer
    isDataDefinition As Boolean
End Type

' Variables globales
Private Cache(0 To CACHE_SIZE - 1) As cacheLine
Private Cache2Way(0 To CACHE_2WAY_SETS - 1, 0 To CACHE_2WAY_BLOCKS_PER_SET - 1) As cacheBlock2Way
Private RAM(0 To RAM_SIZE - 1) As Byte
Private Program() As AssemblyInstruction
Private SymbolTable As Object
Private DataSectionStart As Long
Private TextSectionStart As Long
Private CurrentInstructionIndex As Integer
Private StartInstructionIndex As Integer
Private CacheHits As Integer
Private CacheMisses As Integer
Private TotalAccesses As Integer
Private LastAccessedCacheLine As Integer
Private LastAccessedRAMAddress As Integer
Private ReplacementCount As Integer
Private AccessCounter As Long
Private LogEntries As Collection

' ===================================================================================
' ============================ INICIALIZACIÓN Y CONTROL =============================
' ===================================================================================

' Punto de entrada principal para configurar el simulador
Public Sub InitializeCacheSimulator()
    ' Prepara las hojas de Excel
    Application.ScreenUpdating = False
    CreateSampleProgramSheet
    DrawCacheGrid
    DrawRAMGridCache
    InitializeStatisticsSheet
    InitializeLogSheet
    InitializeStateSheet
    
    ' Inicializa el estado del simulador
    ResetCacheSimulator
    Application.ScreenUpdating = True
    
    MsgBox "Simulador de Caché inicializado. Programa cargado y listo para ejecutar."
End Sub

' Reinicia el simulador a su estado inicial
Public Sub ResetCacheSimulator()
    Application.ScreenUpdating = False
    
    ' Limpiar RAM y caché
    ClearRAM
    ClearCache
    ClearCache2Way
    
    ' Reiniciar variables de estado
    CurrentInstructionIndex = -1
    StartInstructionIndex = -1
    LastAccessedCacheLine = -1
    LastAccessedRAMAddress = -1
    CacheHits = 0
    CacheMisses = 0
    TotalAccesses = 0
    ReplacementCount = 0
    AccessCounter = 0
    Set SymbolTable = CreateObject("Scripting.Dictionary")
    Set LogEntries = New Collection
    
    ' Leer, parsear y cargar el programa
    ReadAndParseNASM
    LoadNASMProgramIntoRAM
    
    ' Buscar el punto de inicio de la ejecución (_start)
    If SymbolTable.Exists("_start") Then
        Dim startAddr As Long
        startAddr = SymbolTable("_start")
        Dim i As Integer
        For i = 0 To UBound(Program)
            If Program(i).address = startAddr And Not Program(i).isDataDefinition Then
                StartInstructionIndex = i
                CurrentInstructionIndex = i
                Exit For
            End If
        Next i
    End If
    
    If StartInstructionIndex = -1 Then
        MsgBox "Advertencia: No se encontró la etiqueta '_start' en la sección .text. La ejecución no puede comenzar.", vbExclamation
    End If
    
    ' Actualizar visualización
    UpdateRAMDisplayCache
    UpdateCacheDisplay
    UpdateCache2WayDisplay
    UpdateStatisticsSheet
    UpdateCacheStatus "LISTO", "---", "Presione 'Siguiente' o 'Ejecutar Todo'"
    
    Application.ScreenUpdating = True
End Sub

' Limpia solo la caché, manteniendo la RAM y el programa
Public Sub ClearCacheOnly()
    Application.ScreenUpdating = False
    ClearCache
    ClearCache2Way
    LastAccessedCacheLine = -1
    UpdateCacheDisplay
    UpdateCache2WayDisplay
    UpdateCacheStats
    UpdateStatisticsSheet
    UpdateCacheStatus "LIMPIO", "---", "Caché vaciada. Estadísticas reiniciadas."
    Application.ScreenUpdating = True
    MsgBox "Caché limpiada. Estadísticas reiniciadas."
End Sub

' ===================================================================================
' =========================== EJECUCIÓN DEL PROGRAMA ================================
' ===================================================================================

' Ejecuta la siguiente instrucción del programa
Public Sub ExecuteNextInstructionCache()
    If CurrentInstructionIndex = -1 Or CurrentInstructionIndex > UBound(Program) Then
        UpdateCacheStatus "COMPLETADO", "---", "El programa ha finalizado."
        MsgBox "Programa completado."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Limpiar resaltados de la ejecución anterior
    LastAccessedCacheLine = -1
    LastAccessedRAMAddress = -1
    UpdateRAMDisplayCache
    
    Dim instruction As AssemblyInstruction
    instruction = Program(CurrentInstructionIndex)
    
    ' Saltar directivas o definiciones de datos
    If instruction.isDataDefinition Or LCase(instruction.Opcode) = "section" Or LCase(instruction.Opcode) = "global" Then
        CurrentInstructionIndex = CurrentInstructionIndex + 1
        ExecuteNextInstructionCache
        Exit Sub
    End If
    
    ' Simular accesos a memoria para la instrucción actual
    SimulateInstructionCacheAccess instruction
    
    ' Actualizar estado visual
    UpdateExecutionStatusCache instruction
    
    ' Avanzar al siguiente índice
    CurrentInstructionIndex = CurrentInstructionIndex + 1
    
    If CurrentInstructionIndex > UBound(Program) Then
        UpdateCacheStatus "COMPLETADO", "---", "Programa finalizado."
    End If
    
    Application.ScreenUpdating = True
End Sub

' Ejecuta el programa completo de una vez
Public Sub ExecuteFullProgramCache()
    If CurrentInstructionIndex = -1 Or CurrentInstructionIndex > UBound(Program) Then
        MsgBox "El programa ya ha finalizado. Por favor, reinicie el simulador.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Do While CurrentInstructionIndex <= UBound(Program)
        ' Avanzar por las líneas que no son instrucciones ejecutables
        Do While Program(CurrentInstructionIndex).isDataDefinition Or LCase(Program(CurrentInstructionIndex).Opcode) = "section" Or LCase(Program(CurrentInstructionIndex).Opcode) = "global"
            CurrentInstructionIndex = CurrentInstructionIndex + 1
            If CurrentInstructionIndex > UBound(Program) Then Exit Do
        Loop
        If CurrentInstructionIndex > UBound(Program) Then Exit Do
        
        Dim instruction As AssemblyInstruction
        instruction = Program(CurrentInstructionIndex)
        
        SimulateInstructionCacheAccess instruction
        UpdateExecutionStatusCache instruction
        
        CurrentInstructionIndex = CurrentInstructionIndex + 1
        
        ' Pequeña pausa para visualización
        DoEvents
    Loop
    
    UpdateCacheStatus "COMPLETADO", "---", "Programa finalizado."
    UpdateStatisticsSheet
    Application.ScreenUpdating = True
    MsgBox "Ejecución completa."
End Sub

' Simula los accesos a memoria que realizaría una instrucción
Private Sub SimulateInstructionCacheAccess(instruction As AssemblyInstruction)
    Dim memAddr As Long
    
    ' Acceso para leer la propia instrucción desde la RAM
    MemoryAccess instruction.address, False ' Lectura de la instrucción
    
    ' Detectar y simular accesos a memoria en operandos
    If InStr(instruction.Operand1, "[") > 0 Then
        memAddr = ExtractAddressFromOperand(instruction.Operand1)
        If LCase(instruction.Opcode) = "mov" Then ' MOV [mem], reg es escritura
            MemoryAccess memAddr, True, &HAA ' Escritura con dato de ejemplo (0xAA)
        Else ' ADD [mem], reg; CMP [mem], reg etc. son lectura
            MemoryAccess memAddr, False
        End If
    End If
    
    If InStr(instruction.Operand2, "[") > 0 Then
        memAddr = ExtractAddressFromOperand(instruction.Operand2)
        MemoryAccess memAddr, False ' Siempre es lectura para el segundo operando (ej: MOV reg, [mem])
    End If
End Sub

' ===================================================================================
' ======================== LÓGICA DE CACHÉ Y MEMORIA ================================
' ===================================================================================

' Simula un acceso a una dirección de memoria, gestionando ambas cachés
Private Function MemoryAccess(address As Long, isWrite As Boolean, Optional Data As Byte = 0) As Boolean
    Dim hitDirect As Boolean
    Dim hit2Way As Boolean
    Dim cacheIndex As Integer
    Dim tag As Long
    Dim offset As Integer
    Dim setIndex As Integer ' CAMBIADO de Long a Integer
    Dim blockIndex As Integer
    
    ' Calcular componentes para mapeo directo
    offset = address Mod CACHE_LINE_SIZE
    cacheIndex = (address \ CACHE_LINE_SIZE) Mod CACHE_SIZE
    tag = address \ (CACHE_SIZE * CACHE_LINE_SIZE)
    
    ' Calcular componentes para caché 2-way
    Dim offset2Way As Integer
    Dim tag2Way As Long
    offset2Way = address Mod CACHE_2WAY_BLOCK_SIZE
    setIndex = (address \ CACHE_2WAY_BLOCK_SIZE) Mod CACHE_2WAY_SETS
    tag2Way = address \ (CACHE_2WAY_SETS * CACHE_2WAY_BLOCK_SIZE)
    
    TotalAccesses = TotalAccesses + 1
    AccessCounter = AccessCounter + 1
    LastAccessedRAMAddress = address
    
    ' Acceso a caché de mapeo directo
    If Cache(cacheIndex).Valid And Cache(cacheIndex).tag = tag Then
        hitDirect = True
        CacheHits = CacheHits + 1
        Cache(cacheIndex).accessCount = Cache(cacheIndex).accessCount + 1
        LastAccessedCacheLine = cacheIndex
        AddLogEntry "HIT", "Directo", setIndex, cacheIndex, "Acceso a dirección: 0x" & Hex(address)
        UpdateCacheStatus "HIT", "0x" & Hex(address), "Línea: " & cacheIndex & ", Tag: 0x" & Hex(tag)
    Else
        hitDirect = False
        CacheMisses = CacheMisses + 1
        LoadCacheLineFromRAM cacheIndex, address, tag
        LastAccessedCacheLine = cacheIndex
        AddLogEntry "MISS", "Directo", setIndex, cacheIndex, "Carga bloque desde RAM: 0x" & Hex(address)
        UpdateCacheStatus "MISS", "0x" & Hex(address), "Línea: " & cacheIndex & " cargada desde RAM"
    End If
    
    ' Acceso a caché 2-way
    hit2Way = AccessCache2Way(address, isWrite, Data)
    
    ' Operación de escritura (Write-Through)
    If isWrite Then
        RAM(address) = Data
        If hitDirect Then
            Cache(cacheIndex).Data(offset) = Data
        End If
        If hit2Way Then
            ' Ya se actualizó en AccessCache2Way
        End If
        UpdateRAMDisplayCache
    End If
    
    ' Actualizar visualizaciones
    UpdateCacheDisplay
    UpdateCache2WayDisplay
    UpdateStatisticsSheet
    
    MemoryAccess = hitDirect
End Function

' Acceso a la caché 2-way con política LRU
Private Function AccessCache2Way(address As Long, isWrite As Boolean, Optional Data As Byte = 0) As Boolean
    Dim setIndex As Integer ' CAMBIADO de Long a Integer
    Dim tag As Long, offset As Integer
    Dim i As Integer, lruIndex As Integer
    Dim hit As Boolean
    
    offset = address Mod CACHE_2WAY_BLOCK_SIZE
    setIndex = (address \ CACHE_2WAY_BLOCK_SIZE) Mod CACHE_2WAY_SETS
    tag = address \ (CACHE_2WAY_SETS * CACHE_2WAY_BLOCK_SIZE)
    
    ' Buscar en el conjunto
    hit = False
    For i = 0 To CACHE_2WAY_BLOCKS_PER_SET - 1
        If Cache2Way(setIndex, i).Valid And Cache2Way(setIndex, i).tag = tag Then
            ' HIT
            hit = True
            Cache2Way(setIndex, i).LastUsed = AccessCounter
            Cache2Way(setIndex, i).accessCount = Cache2Way(setIndex, i).accessCount + 1
            If isWrite Then
                Cache2Way(setIndex, i).Data(offset) = Data
            End If
            Exit For
        End If
    Next i
    
    If Not hit Then
        ' MISS - Buscar bloque para reemplazar (LRU)
        lruIndex = FindLRUBlock(setIndex)
        
        ' Si el bloque era válido, incrementar contador de reemplazos
        If Cache2Way(setIndex, lruIndex).Valid Then
            ReplacementCount = ReplacementCount + 1
            AddLogEntry "REEMPLAZO", "2-Way", setIndex, lruIndex, "Reemplazo LRU en conjunto " & setIndex
        End If
        
        ' Cargar nuevo bloque
        LoadCache2WayBlock setIndex, lruIndex, address, tag
        Cache2Way(setIndex, lruIndex).LastUsed = AccessCounter
        
        If isWrite Then
            Cache2Way(setIndex, lruIndex).Data(offset) = Data
        End If
        
        AddLogEntry "MISS", "2-Way", setIndex, lruIndex, "Carga bloque desde RAM: 0x" & Hex(address)
    Else
        AddLogEntry "HIT", "2-Way", setIndex, i, "Acceso a dirección: 0x" & Hex(address)
    End If
    
    AccessCache2Way = hit
End Function

' Encuentra el bloque LRU en un conjunto
Private Function FindLRUBlock(setIndex As Integer) As Integer
    Dim i As Integer, minUsage As Long, lruIndex As Integer
    minUsage = AccessCounter + 1
    lruIndex = 0
    
    For i = 0 To CACHE_2WAY_BLOCKS_PER_SET - 1
        If Not Cache2Way(setIndex, i).Valid Then
            FindLRUBlock = i
            Exit Function
        End If
        If Cache2Way(setIndex, i).LastUsed < minUsage Then
            minUsage = Cache2Way(setIndex, i).LastUsed
            lruIndex = i
        End If
    Next i
    
    FindLRUBlock = lruIndex
End Function

' Carga bloque en caché 2-way
Private Sub LoadCache2WayBlock(setIndex As Integer, blockIndex As Integer, address As Long, tag As Long)
    Dim baseAddress As Long
    Dim i As Integer
    
    baseAddress = (address \ CACHE_2WAY_BLOCK_SIZE) * CACHE_2WAY_BLOCK_SIZE
    
    With Cache2Way(setIndex, blockIndex)
        .Valid = True
        .tag = tag
        .address = baseAddress
        .accessCount = 1
        
        For i = 0 To CACHE_2WAY_BLOCK_SIZE - 1
            If baseAddress + i < RAM_SIZE Then
                .Data(i) = RAM(baseAddress + i)
            Else
                .Data(i) = 0
            End If
        Next i
    End With
End Sub

' Carga un bloque de memoria de la RAM a una línea específica de la caché
Private Sub LoadCacheLineFromRAM(cacheIndex As Integer, address As Long, tag As Long)
    Dim i As Integer
    Dim baseAddress As Long
    
    ' Calcular la dirección de inicio del bloque en RAM
    baseAddress = (address \ CACHE_LINE_SIZE) * CACHE_LINE_SIZE
    
    With Cache(cacheIndex)
        .Valid = True
        .tag = tag
        .address = baseAddress ' Guardamos la dirección base del bloque
        .accessCount = 1 ' Es el primer acceso a esta nueva línea
        
        ' Copiar los datos del bloque desde la RAM a la línea de caché
        For i = 0 To CACHE_LINE_SIZE - 1
            If baseAddress + i < RAM_SIZE Then
                .Data(i) = RAM(baseAddress + i)
            Else
                .Data(i) = 0 ' Fuera de los límites de la RAM
            End If
        Next i
    End With
End Sub

' Permite al usuario realizar un acceso manual a memoria
Public Sub ManualMemoryAccess()
    Dim addrStr As String
    Dim addr As Long
    Dim accessType As String
    
    addrStr = InputBox("Ingrese la dirección de memoria (en decimal o hexadecimal con '0x'):", "Acceso Manual a Memoria")
    If addrStr = "" Then Exit Sub
    
    On Error Resume Next
    If LCase(left(addrStr, 2)) = "0x" Then
        addr = Application.WorksheetFunction.Hex2Dec(Mid(addrStr, 3))
    Else
        addr = CLng(addrStr)
    End If
    If Err.Number <> 0 Then
        MsgBox "Dirección inválida.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    If addr < 0 Or addr >= RAM_SIZE Then
        MsgBox "La dirección está fuera del rango de la RAM (0-" & RAM_SIZE - 1 & ").", vbCritical
        Exit Sub
    End If
    
    accessType = InputBox("Ingrese el tipo de acceso ('R' para Lectura, 'W' para Escritura):", "Tipo de Acceso")
    If UCase(accessType) = "R" Then
        MemoryAccess addr, False
    ElseIf UCase(accessType) = "W" Then
        MemoryAccess addr, True, &HFF ' Escribir un valor de ejemplo
    Else
        MsgBox "Tipo de acceso no válido.", vbInformation
    End If
End Sub

' ===================================================================================
' ===================== LECTURA Y PARSEO DE CÓDIGO NASM =============================
' ===================================================================================

' Orquesta el proceso de lectura y parseo en dos pasadas
Private Sub ReadAndParseNASM()
    Dim ws As Worksheet
    Set ws = Worksheets("ProgramaNASM")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No se encontró programa en la hoja 'ProgramaNASM'.", vbCritical
        Exit Sub
    End If

    ' Contar líneas válidas para dimensionar el array
    Dim validLines() As String
    Dim validLineCount As Integer
    validLineCount = 0
    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) <> "" Then
            validLineCount = validLineCount + 1
            ReDim Preserve validLines(1 To validLineCount)
            validLines(validLineCount) = Trim(ws.Cells(i, 1).value)
        End If
    Next i

    If validLineCount = 0 Then
        MsgBox "No se encontraron líneas de código válidas.", vbCritical
        Exit Sub
    End If

    ReDim Program(0 To validLineCount - 1)
    
    DataSectionStart = &H0
    TextSectionStart = &H80

    ' --- PRIMERA PASADA: Construir la Tabla de Símbolos ---
    ParsePassOne validLines
    
    ' --- SEGUNDA PASADA: Parsear instrucciones y resolver operandos ---
    ParsePassTwo validLines
End Sub

' Primera pasada: encuentra etiquetas y las añade a la tabla de símbolos
Private Sub ParsePassOne(lines() As String)
    Dim currentAddress As Long
    Dim currentSection As String
    currentSection = ".data" ' Asumir .data por defecto
    currentAddress = DataSectionStart
    Dim i As Integer

    For i = LBound(lines) To UBound(lines)
        Dim line As String, cleanLine As String
        line = lines(i)
        cleanLine = Trim(LCase(Split(line, ";")(0))) ' Ignorar comentarios y espacios

        If InStr(cleanLine, "section .text") > 0 Then
            currentSection = ".text"
            currentAddress = TextSectionStart
        ElseIf InStr(cleanLine, "section .data") > 0 Then
            currentSection = ".data"
            currentAddress = DataSectionStart
        ElseIf cleanLine <> "" Then ' Solo procesar líneas no vacías
            Dim parts() As String
            parts = Split(Trim(line), " ")
            
            ' Es una etiqueta (ej: _start:)
            If Right(parts(0), 1) = ":" Then
                Dim labelName As String
                labelName = left(parts(0), Len(parts(0)) - 1)
                SymbolTable(labelName) = currentAddress
                ' Las etiquetas no consumen espacio de memoria
            ElseIf UBound(parts) >= 1 Then
                ' Es una definición de datos (dd, db, dw)
                If LCase(parts(1)) = "dd" Or LCase(parts(1)) = "db" Or LCase(parts(1)) = "dw" Then
                    SymbolTable(parts(0)) = currentAddress
                    ' Avanzar la dirección según el tamaño del dato
                    Select Case LCase(parts(1))
                        Case "dd": currentAddress = currentAddress + 4
                        Case "dw": currentAddress = currentAddress + 2
                        Case "db": currentAddress = currentAddress + 1
                    End Select
                ElseIf currentSection = ".text" And LCase(parts(0)) <> "global" Then
                    ' Es una instrucción, avanzar 4 bytes (simplificación)
                    currentAddress = currentAddress + 4
                End If
            End If
        End If
    Next i
End Sub

' Segunda pasada: parsea cada línea en la estructura de instrucción
Private Sub ParsePassTwo(lines() As String)
    Dim currentAddress As Long
    Dim currentSection As String
    currentSection = ".data" ' Asumir .data por defecto
    currentAddress = DataSectionStart
    Dim i As Integer

    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = lines(i)
        
        With Program(i - 1)
            .OriginalLine = line
            
            Dim cleanLine As String
            cleanLine = LCase(Split(line, ";")(0))
            
            If InStr(cleanLine, "section .text") > 0 Then
                currentSection = ".text"
                currentAddress = TextSectionStart
                .Opcode = "section"
                .address = currentAddress
            ElseIf InStr(cleanLine, "section .data") > 0 Then
                currentSection = ".data"
                currentAddress = DataSectionStart
                .Opcode = "section"
                .address = currentAddress
            Else
                .address = currentAddress
                ParseInstructionLine line, i - 1
                
                ' Avanzar dirección si no es una etiqueta
                If Not Right(.Opcode, 1) = ":" And .Opcode <> "global" Then
                    currentAddress = currentAddress + .Length
                End If
            End If
        End With
    Next i
End Sub

' Parsea una línea de instrucción individual
Private Sub ParseInstructionLine(line As String, index As Integer)
    Dim mainParts() As String, parts() As String, commentPart As String
    
    ' Separar comentarios
    If InStr(line, ";") > 0 Then
        commentPart = Split(line, ";")(0)
    Else
        commentPart = line
    End If
    
    ' Validar que la línea no esté vacía después de quitar comentarios
    If Trim(commentPart) = "" Then
        With Program(index)
            .Opcode = ""
            .Operand1 = ""
            .Operand2 = ""
            .Length = 0
            .isDataDefinition = False
        End With
        Exit Sub
    End If
    
    ' Separar operandos por coma
    mainParts = Split(commentPart, ",")
    
    ' Separar opcode del primer operando (usando espacio como separador)
    parts = Split(Trim(mainParts(0)), " ")
    
    With Program(index)
        ' Verificar que parts tenga al menos un elemento
        If UBound(parts) >= 0 Then
            .Opcode = Trim(parts(0))
        Else
            .Opcode = ""
        End If
        
        ' Verificar y asignar operandos de manera segura
        If UBound(parts) >= 1 Then
            .Operand1 = Trim(parts(1))
        Else
            .Operand1 = ""
        End If
        
        If UBound(mainParts) >= 1 Then
            .Operand2 = Trim(mainParts(1))
        Else
            .Operand2 = ""
        End If
        
        ' Asignar longitud y tipo de instrucción
        If .Opcode <> "" And LCase(.Opcode) <> "section" And LCase(.Opcode) <> "global" And Right(.Opcode, 1) <> ":" Then
            .Length = 4 ' Longitud estándar para instrucciones
            .bytes = GenerateSimpleBytes(.Opcode)
            
            ' Verificar si es una definición de datos
            If LCase(.Opcode) = "dd" Or LCase(.Opcode) = "dw" Or LCase(.Opcode) = "db" Then
                .isDataDefinition = True
                Select Case LCase(.Opcode)
                    Case "dd": .Length = 4
                    Case "dw": .Length = 2
                    Case "db": .Length = 1
                End Select
            Else
                .isDataDefinition = False
            End If
        Else
            .Length = 0
            .isDataDefinition = False
        End If
    End With
End Sub

' Carga los bytes del programa en la RAM simulada
Private Sub LoadNASMProgramIntoRAM()
    Dim i As Integer, j As Integer
    Dim bytes() As String
    Dim byteValue As Byte
    
    For i = 0 To UBound(Program)
        If Not Program(i).isDataDefinition And Program(i).bytes <> "" Then
            bytes = Split(Program(i).bytes, " ")
            For j = 0 To UBound(bytes)
                If Program(i).address + j < RAM_SIZE Then
                    byteValue = CInt("&H" & bytes(j))
                    RAM(Program(i).address + j) = byteValue
                End If
            Next j
        End If
    Next i
End Sub

' Extrae la dirección de un operando como [num1] o [128]
Private Function ExtractAddressFromOperand(operand As String) As Long
    Dim cleanOperand As String
    cleanOperand = Replace(Replace(operand, "[", ""), "]", "")
    
    ' Si es una variable, buscar en la tabla de símbolos
    If SymbolTable.Exists(cleanOperand) Then
        ExtractAddressFromOperand = SymbolTable(cleanOperand)
    Else ' Si no, asumir que es una dirección numérica
        On Error Resume Next
        ExtractAddressFromOperand = CLng(cleanOperand)
        If Err.Number <> 0 Then ExtractAddressFromOperand = -1 ' Dirección inválida
        On Error GoTo 0
    End If
End Function

' Genera bytes de máquina falsos para la simulación visual
Private Function GenerateSimpleBytes(Opcode As String) As String
    Select Case LCase(Opcode)
        Case "mov": GenerateSimpleBytes = "B8 12 34 56"
        Case "add": GenerateSimpleBytes = "01 C0 90 90"
        Case "xor": GenerateSimpleBytes = "31 DB 90 90"
        Case "int": GenerateSimpleBytes = "CD 80 90 90"
        Case Else: GenerateSimpleBytes = "90 90 90 90" ' NOP
    End Select
End Function

' ===================================================================================
' ======================== HOJAS DE VISUALIZACIÓN ===================================
' ===================================================================================

' Inicializa hoja de Estadísticas
Private Sub InitializeStatisticsSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("Estadísticas")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "Estadísticas"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws
        .Range("A1").value = "ESTADÍSTICAS DE CACHÉ"
        .Range("A1:D1").Merge
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A3").value = "Total Accesos:"
        .Range("B3").value = "0"
        
        .Range("A4").value = "Hits:"
        .Range("B4").value = "0"
        .Range("B4").Interior.Color = RGB(180, 255, 180)
        
        .Range("A5").value = "Misses:"
        .Range("B5").value = "0"
        .Range("B5").Interior.Color = RGB(255, 180, 180)
        
        .Range("A6").value = "Reemplazos:"
        .Range("B6").value = "0"
        .Range("B6").Interior.Color = RGB(255, 220, 180)
        
        .Range("A8").value = "Configuración:"
        .Range("A8").Font.Bold = True
        
        .Range("A9").value = "Tamaño Caché:"
        .Range("B9").value = "1 KB"
        
        .Range("A10").value = "Tamaño Bloque:"
        .Range("B10").value = "16 bytes"
        
        .Range("A11").value = "Asociatividad:"
        .Range("B11").value = "2-way"
        
        .Range("A12").value = "Número de Conjuntos:"
        .Range("B12").value = "32"
        
        .Columns("A:B").AutoFit
    End With
End Sub

' Inicializa hoja de Log
Private Sub InitializeLogSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("LogCaché")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "LogCaché"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws
        .Range("A1:H1").Merge
        .Range("A1").value = "LOG DE CACHÉ - Registro de Accesos"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        
        Dim headers As Variant
        headers = Array("Timestamp", "Direction/Evento", "Tipo", "Set", "Bloque", "Información")
        .Range("A3").Resize(1, UBound(headers) + 1).value = headers
        
        With .Range("A3").Resize(1, UBound(headers) + 1)
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
            .HorizontalAlignment = xlCenter
        End With
        
        .Columns("A:F").AutoFit
        .Columns("F").ColumnWidth = 40
    End With
End Sub

' Inicializa hoja de Estado
Private Sub InitializeStateSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("EstadoCaché")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "EstadoCaché"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws
        .Range("A1:I1").Merge
        .Range("A1").value = "ESTADO DE CACHÉ 2-WAY SET ASSOCIATIVE"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        
        Dim headers As Variant
        headers = Array("Conjunto", "Bloque", "Válido", "Tag", "Dirección", "Datos", "Último Uso", "LRU Rank")
        .Range("A2").Resize(1, UBound(headers) + 1).value = headers
        
        With .Range("A2").Resize(1, UBound(headers) + 1)
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Llenar estructura base
        Dim i As Integer, j As Integer, row As Integer
        row = 3
        
        For i = 0 To CACHE_2WAY_SETS - 1
            For j = 0 To CACHE_2WAY_BLOCKS_PER_SET - 1
                .Cells(row, 1).value = i
                .Cells(row, 2).value = j
                .Cells(row, 3).value = "No"
                .Cells(row, 4).value = "N/A"
                .Cells(row, 5).value = ""
                .Cells(row, 6).value = "N/A"
                .Cells(row, 7).value = ""
                .Cells(row, 8).value = ""
                row = row + 1
            Next j
            row = row + 1 ' Espacio entre conjuntos
        Next i
        
        .Columns("A:I").AutoFit
    End With
End Sub

' Actualiza hoja de Estadísticas
Private Sub UpdateStatisticsSheet()
    Dim ws As Worksheet
    Set ws = Worksheets("Estadísticas")
    
    With ws
        .Range("B3").value = TotalAccesses
        .Range("B4").value = CacheHits
        .Range("B5").value = CacheMisses
        .Range("B6").value = ReplacementCount
    End With
End Sub

' Añade entrada al log
Private Sub AddLogEntry(eventType As String, cacheType As String, setIndex As Integer, blockIndex As Integer, info As String)
    Dim ws As Worksheet
    Set ws = Worksheets("LogCaché")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row + 1
    
    If lastRow < 4 Then lastRow = 4
    
    With ws
        .Cells(lastRow, 1).value = Format(Now, "hh:mm:ss")
        .Cells(lastRow, 2).value = eventType
        .Cells(lastRow, 3).value = cacheType
        .Cells(lastRow, 4).value = setIndex
        .Cells(lastRow, 5).value = blockIndex
        .Cells(lastRow, 6).value = info
        
        ' Colores según el tipo de evento
        Select Case eventType
            Case "HIT"
                .Range(.Cells(lastRow, 1), .Cells(lastRow, 6)).Interior.Color = RGB(220, 255, 220)
            Case "MISS"
                .Range(.Cells(lastRow, 1), .Cells(lastRow, 6)).Interior.Color = RGB(255, 220, 220)
            Case "REEMPLAZO"
                .Range(.Cells(lastRow, 1), .Cells(lastRow, 6)).Interior.Color = RGB(255, 255, 220)
        End Select
    End With
    
    ' Autoajustar columnas
    ws.Columns("A:F").AutoFit
End Sub

' Actualiza visualización de caché 2-way
Private Sub UpdateCache2WayDisplay()
    Dim ws As Worksheet
    Set ws = Worksheets("EstadoCaché")
    
    Dim row As Integer, i As Integer, j As Integer
    row = 3
    
    For i = 0 To CACHE_2WAY_SETS - 1
        For j = 0 To CACHE_2WAY_BLOCKS_PER_SET - 1
            With Cache2Way(i, j)
                ws.Cells(row, 1).value = i
                ws.Cells(row, 2).value = j
                ws.Cells(row, 3).value = IIf(.Valid, "Sí", "No")
                ws.Cells(row, 4).value = IIf(.Valid, "0x" & Hex(.tag), "N/A")
                ws.Cells(row, 5).value = IIf(.Valid, "0x" & Format(Hex(.address), "000"), "")
                
                ' Datos (primeros 4 bytes como ejemplo)
                If .Valid Then
                    ws.Cells(row, 6).value = Format(Hex(.Data(0)), "00") & " " & Format(Hex(.Data(1)), "00") & " ..."
                Else
                    ws.Cells(row, 6).value = "N/A"
                End If
                
                ws.Cells(row, 7).value = IIf(.Valid, .LastUsed, "")
                ws.Cells(row, 8).value = CalculateLRURank(i, j)
            End With
            
            ' Color de fondo
            If Cache2Way(i, j).Valid Then
                ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Interior.Color = RGB(220, 255, 220)
            Else
                ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Interior.Color = RGB(255, 220, 220)
            End If
            
            row = row + 1
        Next j
        row = row + 1 ' Espacio entre conjuntos
    Next i
End Sub

' Calcula el ranking LRU para un bloque
Private Function CalculateLRURank(setIndex As Integer, blockIndex As Integer) As String
    If Not Cache2Way(setIndex, blockIndex).Valid Then
        CalculateLRURank = "N/A"
        Exit Function
    End If
    
    Dim i As Integer, rank As Integer
    rank = 1
    
    For i = 0 To CACHE_2WAY_BLOCKS_PER_SET - 1
        If i <> blockIndex And Cache2Way(setIndex, i).Valid Then
            If Cache2Way(setIndex, i).LastUsed < Cache2Way(setIndex, blockIndex).LastUsed Then
                rank = rank + 1
            End If
        End If
    Next i
    
    CalculateLRURank = rank & "/" & CACHE_2WAY_BLOCKS_PER_SET
End Function

' ===================================================================================
' ================== DIBUJO Y ACTUALIZACIÓN DE LA INTERFAZ ==========================
' ===================================================================================

' Dibuja la cuadrícula de la caché y los controles
Private Sub DrawCacheGrid()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("Cache")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "Cache"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Activate
    
    With ws
        .Range("B1").value = "MEMORIA CACHÉ - Mapeo Directo"
        .Range("B1:G1").Merge
        .Range("B1").Font.Bold = True
        .Range("B1").Font.Size = 16
        
        .Range("I1").value = "ESTADÍSTICAS DE CACHÉ": .Range("I1:K1").Merge: .Range("I1").Font.Bold = True
        .Range("I2").value = "Total Accesos:": .Range("J2").value = 0
        .Range("I3").value = "Cache Hits:": .Range("J3").value = 0: .Range("J3").Interior.Color = RGB(180, 255, 180)
        .Range("I4").value = "Cache Misses:": .Range("J4").value = 0: .Range("J4").Interior.Color = RGB(255, 180, 180)
        .Range("I5").value = "Hit Rate:": .Range("J5").value = "0%"
        
        Dim headers As Variant
        headers = Array("Línea", "Válido", "Tag (Hex)", "Dirección Bloque", "Datos (Hex)", "Accesos")
        .Range("A3").Resize(1, UBound(headers) + 1).value = headers
        With .Range("A3").Resize(1, UBound(headers) + 1)
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
            .HorizontalAlignment = xlCenter
        End With

        Dim i As Integer
        For i = 0 To CACHE_SIZE - 1
            .Cells(i + 4, 1).value = i
        Next i
        
        With .Range("A4").Resize(CACHE_SIZE, UBound(headers) + 1)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        
        .Columns("A:G").AutoFit
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 20
    End With
    
    CreateCacheControlButtons
End Sub

' Dibuja la cuadrícula de la RAM y el panel de estado
Private Sub DrawRAMGridCache()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("RAM")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "RAM"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws
        .Range("B1").value = "MEMORIA RAM": .Range("B1:J1").Merge: .Range("B1").Font.Bold = True: .Range("B1").Font.Size = 16
        
        .Range("S2").value = "ESTADO DEL SIMULADOR": .Range("S2:U2").Merge: .Range("S2").Font.Bold = True
        .Range("S3").value = "Última Instrucción:"
        .Range("S4").value = "Último Acceso:"
        .Range("S5").value = "Dirección:"
        .Range("S6").value = "Detalles:"
        
        .Range("T3:T6").value = "---"
        .Range("T3:T6").HorizontalAlignment = xlLeft
        .Range("S3:S6").Font.Bold = True
        .Columns("S").AutoFit
        .Columns("T").ColumnWidth = 30
        
        ' Encabezados de columnas (0-F)
        Dim i As Integer
        For i = 0 To COLS - 1
            .Cells(8, i + 2).value = Hex(i)
        Next i
        With .Range(.Cells(8, 2), .Cells(8, COLS + 1))
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(220, 220, 220)
        End With

        ' Encabezados de filas y celdas
        Dim r As Integer, addr As Long
        For r = 0 To ROWS - 1
            addr = r * COLS
            .Cells(r + 9, 1).value = "0x" & Format(Hex(addr), "00")
            For i = 0 To COLS - 1
                .Cells(r + 9, i + 2).value = "00"
            Next i
        Next r
        
        With .Range("A9").Resize(ROWS, 1)
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
        End With
        With .Range("B9").Resize(ROWS, COLS)
            .HorizontalAlignment = xlCenter
            .Font.Name = "Courier New"
            .Borders.LineStyle = xlContinuous
        End With
    End With
End Sub

' Actualiza la visualización de la RAM
Private Sub UpdateRAMDisplayCache()
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    Dim r As Integer, c As Integer, addr As Long
    
    For r = 0 To ROWS - 1
        For c = 0 To COLS - 1
            addr = r * COLS + c
            With ws.Cells(r + 9, c + 2)
                .value = Format(Hex(RAM(addr)), "00")
                
                ' Color de fondo por sección
                If addr >= TextSectionStart Then
                    .Interior.Color = RGB(220, 255, 220) ' Verde para .text
                Else
                    .Interior.Color = RGB(220, 220, 255) ' Azul para .data
                End If
                .Font.Bold = False
            End With
        Next c
    Next r
    
    ' Resaltar el último acceso a RAM
    If LastAccessedRAMAddress <> -1 Then
        r = (LastAccessedRAMAddress \ COLS)
        c = (LastAccessedRAMAddress Mod COLS)
        With ws.Cells(r + 9, c + 2)
            .Interior.Color = RGB(255, 255, 0) ' Amarillo
            .Font.Bold = True
        End With
    End If
End Sub

' Actualiza la visualización de la caché
Private Sub UpdateCacheDisplay()
    Dim ws As Worksheet
    Set ws = Worksheets("Cache")
    Dim i As Integer, j As Integer
    
    For i = 0 To CACHE_SIZE - 1
        With Cache(i)
            ' Válido
            ws.Cells(i + 4, 2).value = IIf(.Valid, "Sí", "No")
            ' Tag
            ws.Cells(i + 4, 3).value = "0x" & Hex(.tag)
            ' Dirección y Datos
            If .Valid Then
                ws.Cells(i + 4, 4).value = "0x" & Format(Hex(.address), "00")
                Dim dataStr As String
                dataStr = ""
                For j = 0 To CACHE_LINE_SIZE - 1
                    dataStr = dataStr & Format(Hex(.Data(j)), "00") & " "
                Next j
                ws.Cells(i + 4, 5).value = Trim(dataStr)
            Else
                ws.Cells(i + 4, 4).value = "---"
                ws.Cells(i + 4, 5).value = "00 00 00 00"
            End If
            ' Accesos
            ws.Cells(i + 4, 6).value = .accessCount
        End With
        
        ' Color de fondo base (sin el resaltado de acceso)
        If Cache(i).Valid Then
            ws.Range("B" & i + 4 & ":F" & i + 4).Interior.Color = RGB(220, 255, 220) ' Verde claro
        Else
            ws.Range("B" & i + 4 & ":F" & i + 4).Interior.Color = RGB(255, 220, 220) ' Rojo claro
        End If
    Next i
    
    ' Resaltar la última línea accedida
    If LastAccessedCacheLine <> -1 Then
        ws.Range("A" & LastAccessedCacheLine + 4 & ":F" & LastAccessedCacheLine + 4).Interior.Color = RGB(255, 255, 0) ' Amarillo
    End If
    
    UpdateCacheStats
End Sub

' Actualiza el panel de estado general en la hoja RAM
Private Sub UpdateCacheStatus(status As String, address As String, details As String)
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    
    ws.Range("T4").value = status
    ws.Range("T5").value = address
    ws.Range("T6").value = details
    
    ' Color según el estado
    Select Case status
        Case "HIT": ws.Range("T4").Interior.Color = RGB(0, 255, 0)
        Case "MISS": ws.Range("T4").Interior.Color = RGB(255, 0, 0)
        Case "COMPLETADO": ws.Range("T4").Interior.Color = RGB(0, 255, 255)
        Case "LISTO", "LIMPIO": ws.Range("T4").Interior.Color = RGB(200, 200, 255)
        Case Else: ws.Range("T4").Interior.Color = RGB(255, 255, 255)
    End Select
End Sub

' Actualiza el indicador de la instrucción actual en la hoja RAM
Private Sub UpdateExecutionStatusCache(instruction As AssemblyInstruction)
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    ws.Range("T3").value = instruction.OriginalLine
    ws.Range("T3").Interior.Color = RGB(255, 255, 0) ' Amarillo
End Sub

' Actualiza las estadísticas de la caché
Private Sub UpdateCacheStats()
    Dim ws As Worksheet
    Set ws = Worksheets("Cache")
    
    ws.Range("J2").value = TotalAccesses
    ws.Range("J3").value = CacheHits
    ws.Range("J4").value = CacheMisses
    
    If TotalAccesses > 0 Then
        ws.Range("J5").value = Format((CacheHits / TotalAccesses), "0.0%")
    Else
        ws.Range("J5").value = "0%"
    End If
End Sub

' ===================================================================================
' ========================= UTILIDADES Y CONFIGURACIÓN ==============================
' ===================================================================================

' Limpia la memoria RAM
Private Sub ClearRAM()
    Dim i As Long
    For i = 0 To RAM_SIZE - 1
        RAM(i) = 0
    Next i
End Sub

' Limpia la caché y resetea las estadísticas
Private Sub ClearCache()
    Dim i As Integer, j As Integer
    For i = 0 To CACHE_SIZE - 1
        Cache(i).Valid = False
        Cache(i).tag = 0
        Cache(i).address = 0
        Cache(i).accessCount = 0
        For j = 0 To CACHE_LINE_SIZE - 1
            Cache(i).Data(j) = 0
        Next j
    Next i
    
    CacheHits = 0
    CacheMisses = 0
    TotalAccesses = 0
    ReplacementCount = 0
    AccessCounter = 0
End Sub

' Limpia la caché 2-way
Private Sub ClearCache2Way()
    Dim i As Integer, j As Integer
    For i = 0 To CACHE_2WAY_SETS - 1
        For j = 0 To CACHE_2WAY_BLOCKS_PER_SET - 1
            With Cache2Way(i, j)
                .Valid = False
                .tag = 0
                .address = 0
                .LastUsed = 0
                .accessCount = 0
                Dim k As Integer
                For k = 0 To CACHE_2WAY_BLOCK_SIZE - 1
                    .Data(k) = 0
                Next k
            End With
        Next j
    Next i
End Sub

' Crea los botones de control en la hoja de la Caché
Private Sub CreateCacheControlButtons()
    Dim ws As Worksheet
    Set ws = Worksheets("Cache")
    
    ' Limpiar botones existentes para evitar duplicados
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    Dim btn As Button
    Set btn = ws.Buttons.Add(50, 350, 120, 30)
    btn.OnAction = "ExecuteNextInstructionCache"
    btn.Characters.text = "Siguiente Instrucción"
    
    Set btn = ws.Buttons.Add(180, 350, 120, 30)
    btn.OnAction = "ExecuteFullProgramCache"
    btn.Characters.text = "Ejecutar Todo"
    
    Set btn = ws.Buttons.Add(310, 350, 120, 30)
    btn.OnAction = "ResetCacheSimulator"
    btn.Characters.text = "Reiniciar Simulador"
    
    Set btn = ws.Buttons.Add(440, 350, 120, 30)
    btn.OnAction = "ManualMemoryAccess"
    btn.Characters.text = "Acceso Manual"
    
    Set btn = ws.Buttons.Add(570, 350, 100, 30)
    btn.OnAction = "ClearCacheOnly"
    btn.Characters.text = "Limpiar Caché"
End Sub

Private Sub CreateSampleProgramSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("ProgramaNASM")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "ProgramaNASM"
    End If
    
    If ws.Range("A2").value <> "" Then Exit Sub ' No sobrescribir si ya hay algo
    
    ws.Cells.Clear
    ws.Range("A1").value = "Código NASM (Puede editar este código y reiniciar el simulador)"
    ws.Range("A1").Font.Bold = True
    
    ' Escribir el programa línea por línea
    Dim row As Integer
    row = 2
    
    ws.Cells(row, 1).value = "section .data": row = row + 1
    ws.Cells(row, 1).value = "    num1 dd 10": row = row + 1
    ws.Cells(row, 1).value = "    num2 dd 20": row = row + 1
    ws.Cells(row, 1).value = "    result dd 0": row = row + 1
    ws.Cells(row, 1).value = "": row = row + 1
    ws.Cells(row, 1).value = "section .text": row = row + 1
    ws.Cells(row, 1).value = "    global _start": row = row + 1
    ws.Cells(row, 1).value = "": row = row + 1
    ws.Cells(row, 1).value = "_start:": row = row + 1
    ws.Cells(row, 1).value = "    ; Cargar num1 en EAX": row = row + 1
    ws.Cells(row, 1).value = "    mov eax, [num1]": row = row + 1
    ws.Cells(row, 1).value = "    ; Sumar num2 a EAX": row = row + 1
    ws.Cells(row, 1).value = "    add eax, [num2]": row = row + 1
    ws.Cells(row, 1).value = "    ; Guardar resultado": row = row + 1
    ws.Cells(row, 1).value = "    mov [result], eax": row = row + 1
    ws.Cells(row, 1).value = "    ; Acceso a dirección fija para demostrar caché": row = row + 1
    ws.Cells(row, 1).value = "    mov ebx, [128]": row = row + 1
    ws.Cells(row, 1).value = "    ; Salir del programa": row = row + 1
    ws.Cells(row, 1).value = "    mov eax, 1": row = row + 1
    ws.Cells(row, 1).value = "    xor ebx, ebx": row = row + 1
    ws.Cells(row, 1).value = "    int 0x80": row = row + 1
    
    ws.Columns("A").ColumnWidth = 50
    ws.Columns("A").WrapText = True
End Sub

' ===================================================================================
' ========================== FUNCIONES PÚBLICAS =====================================
' ===================================================================================

Public Sub IniciarSimuladorCache()
    InitializeCacheSimulator
End Sub

Public Sub EjecutarProgramaCompletoCache()
    ExecuteFullProgramCache
End Sub
