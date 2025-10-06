Attribute VB_Name = "Pipeline"
' ===== ModuloPipeline =====
Option Explicit

Type PipelineInstruction
    instruction As String
    Opcode As String
    Operand1 As String
    Operand2 As String
    stage As String
    CycleEntered As Long
    CurrentStageCycle As Long
    Color As Long
    Result As String
    Stalled As Boolean
    InstructionNumber As Long
End Type

' Variables globales del Pipeline
Dim Pipeline() As PipelineInstruction
Dim ClockCycle As Long
Dim Instructions() As String
Dim CurrentInstructionIndex As Long
Dim PipelineStages(4) As String
Dim IsPipelineRunning As Boolean
Dim InstructionColors As Collection

' =============================================
' INICIALIZACIÓN PIPELINE
' =============================================

Sub IniciarPipeline()
    InitializePipeline
    CreatePipelineDisplay
    UpdatePipelineDisplay
End Sub

Sub InitializePipeline()
    ' Configurar etapas
    PipelineStages(0) = "IF"
    PipelineStages(1) = "ID"
    PipelineStages(2) = "EX"
    PipelineStages(3) = "MEM"
    PipelineStages(4) = "WB"
    
    ' Inicializar pipeline vacío
    ReDim Pipeline(0 To 4)
    Dim i As Integer
    For i = 0 To 4
        ClearPipelineSlot i
    Next i
    
    ClockCycle = 0
    CurrentInstructionIndex = 0
    Set InstructionColors = New Collection
    IsPipelineRunning = False
    
    ' Cargar instrucciones
    CargarInstruccionesPipeline
End Sub

Sub CargarInstruccionesPipeline()
    ' Cargar desde hoja Programa
    Dim wsPrograma As Worksheet
    Set wsPrograma = ThisWorkbook.Sheets("Programa")
    
    Dim contenido As String
    contenido = Trim(wsPrograma.Range("A6").value)
    
    If contenido = "" Then
        ReDim Instructions(0 To 0)
        Instructions(0) = "NOP"
        Exit Sub
    End If
    
    ' Dividir por saltos de línea
    Dim lineas() As String
    lineas = Split(contenido, vbLf)
    
    If UBound(lineas) = 0 Then
        lineas = Split(contenido, vbCrLf)
    End If
    
    ' Cargar instrucciones
    Dim count As Long
    count = 0
    Dim i As Long
    
    ' Redimensionar array temporal
    ReDim Instructions(0 To UBound(lineas))
    
    For i = 0 To UBound(lineas)
        Dim linea As String
        linea = Trim(lineas(i))
        
        If linea <> "" And left(linea, 1) <> ";" Then
            Instructions(count) = linea
            count = count + 1
        End If
    Next i
    
    ' Redimensionar array final
    If count > 0 Then
        ReDim Preserve Instructions(0 To count - 1)
    Else
        ReDim Instructions(0 To 0)
        Instructions(0) = "NOP"
    End If
End Sub

Function GetInstructionCount() As Long
    On Error GoTo ErrorHandler
    If Not IsArrayInitialized(Instructions) Then
        GetInstructionCount = 0
        Exit Function
    End If
    GetInstructionCount = UBound(Instructions) - LBound(Instructions) + 1
    Exit Function
ErrorHandler:
    GetInstructionCount = 0
End Function

Function IsArrayInitialized(arr As Variant) As Boolean
    On Error GoTo ErrorHandler
    If Not IsArray(arr) Then
        IsArrayInitialized = False
        Exit Function
    End If
    Dim test As Long
    test = UBound(arr)
    IsArrayInitialized = True
    Exit Function
ErrorHandler:
    IsArrayInitialized = False
End Function

Sub ClearPipelineSlot(index As Integer)
    Pipeline(index).instruction = ""
    Pipeline(index).Opcode = ""
    Pipeline(index).Operand1 = ""
    Pipeline(index).Operand2 = ""
    Pipeline(index).stage = ""
    Pipeline(index).CycleEntered = 0
    Pipeline(index).CurrentStageCycle = 0
    Pipeline(index).Color = RGB(255, 255, 255)
    Pipeline(index).Result = ""
    Pipeline(index).Stalled = False
    Pipeline(index).InstructionNumber = 0
End Sub

' =============================================
' SIMULACIÓN PIPELINE
' =============================================

Public Sub EjecutarPipelineCompleto()
    IsPipelineRunning = True
    Dim initialCycle As Long
    initialCycle = ClockCycle
    
    Do While IsPipelineRunning And (ClockCycle - initialCycle) < 50
        ClockCycle = ClockCycle + 1
        AdvancePipeline
        UpdatePipelineDisplay
        DoEvents
        
        If AllInstructionsCompleted() Then
            IsPipelineRunning = False
            MsgBox "? Pipeline completado en " & ClockCycle & " ciclos", vbInformation
            Exit Do
        End If
        
        ' Pausa más corta
        Dim inicioPausa As Single
        inicioPausa = Timer
        Do While Timer < inicioPausa + 0.5
            DoEvents
        Loop
    Loop
    
    IsPipelineRunning = False
End Sub

Public Sub AvanzarCicloPipeline()
    If IsPipelineRunning Then
        Exit Sub
    End If
    
    ClockCycle = ClockCycle + 1
    AdvancePipeline
    UpdatePipelineDisplay
    
    If AllInstructionsCompleted() Then
        MsgBox "? Todas las instrucciones procesadas en " & ClockCycle & " ciclos", vbInformation
    End If
End Sub

Sub AdvancePipeline()
    ' Mover instrucciones de derecha a izquierda (WB -> IF)
    
    ' WB: completar y liberar
    If Pipeline(4).stage = "WB" Then
        Pipeline(4).CurrentStageCycle = Pipeline(4).CurrentStageCycle + 1
        If Pipeline(4).CurrentStageCycle >= 1 Then
            ClearPipelineSlot 4
        End If
    End If
    
    ' MEM -> WB
    If Pipeline(3).stage = "MEM" And Pipeline(4).stage = "" Then
        Pipeline(3).CurrentStageCycle = Pipeline(3).CurrentStageCycle + 1
        If Pipeline(3).CurrentStageCycle >= 1 Then
            Pipeline(4) = Pipeline(3)
            Pipeline(4).stage = "WB"
            Pipeline(4).CurrentStageCycle = 0
            ProcessInstructionInStage 4
            ClearPipelineSlot 3
        End If
    ElseIf Pipeline(3).stage = "MEM" Then
        Pipeline(3).Stalled = True
    End If
    
    ' EX -> MEM
    If Pipeline(2).stage = "EX" And Pipeline(3).stage = "" Then
        Pipeline(2).CurrentStageCycle = Pipeline(2).CurrentStageCycle + 1
        If Pipeline(2).CurrentStageCycle >= 1 Then
            Pipeline(3) = Pipeline(2)
            Pipeline(3).stage = "MEM"
            Pipeline(3).CurrentStageCycle = 0
            Pipeline(3).Stalled = False
            ProcessInstructionInStage 3
            ClearPipelineSlot 2
        End If
    ElseIf Pipeline(2).stage = "EX" Then
        Pipeline(2).Stalled = True
    End If
    
    ' ID -> EX
    If Pipeline(1).stage = "ID" And Pipeline(2).stage = "" Then
        Pipeline(1).CurrentStageCycle = Pipeline(1).CurrentStageCycle + 1
        If Pipeline(1).CurrentStageCycle >= 1 Then
            Pipeline(2) = Pipeline(1)
            Pipeline(2).stage = "EX"
            Pipeline(2).CurrentStageCycle = 0
            Pipeline(2).Stalled = False
            ProcessInstructionInStage 2
            ClearPipelineSlot 1
        End If
    ElseIf Pipeline(1).stage = "ID" Then
        Pipeline(1).Stalled = True
    End If
    
    ' IF -> ID
    If Pipeline(0).stage = "IF" And Pipeline(1).stage = "" Then
        Pipeline(0).CurrentStageCycle = Pipeline(0).CurrentStageCycle + 1
        If Pipeline(0).CurrentStageCycle >= 1 Then
            Pipeline(1) = Pipeline(0)
            Pipeline(1).stage = "ID"
            Pipeline(1).CurrentStageCycle = 0
            Pipeline(1).Stalled = False
            ProcessInstructionInStage 1
            ClearPipelineSlot 0
        End If
    ElseIf Pipeline(0).stage = "IF" Then
        Pipeline(0).Stalled = True
    End If
    
    ' Nueva instrucción -> IF
    If Pipeline(0).stage = "" And CurrentInstructionIndex <= UBound(Instructions) Then
        InsertNewInstruction Instructions(CurrentInstructionIndex), CurrentInstructionIndex + 1
        CurrentInstructionIndex = CurrentInstructionIndex + 1
    End If
End Sub

Sub InsertNewInstruction(instruction As String, instNum As Long)
    Pipeline(0).instruction = instruction
    Pipeline(0).stage = "IF"
    Pipeline(0).CycleEntered = ClockCycle
    Pipeline(0).CurrentStageCycle = 0
    Pipeline(0).Stalled = False
    Pipeline(0).InstructionNumber = instNum
    Pipeline(0).Color = GetInstructionColor(instNum)
    ProcessInstructionInStage 0
End Sub

Sub ProcessInstructionInStage(stageIndex As Integer)
    Select Case Pipeline(stageIndex).stage
        Case "IF"
            Pipeline(stageIndex).Result = "Fetching..."
        Case "ID"
            ParseInstruction stageIndex
            Pipeline(stageIndex).Result = "Decoding"
        Case "EX"
            ExecuteInstruction stageIndex
        Case "MEM"
            Pipeline(stageIndex).Result = "Memory Access"
        Case "WB"
            Pipeline(stageIndex).Result = "Writing Back"
    End Select
End Sub

Sub ParseInstruction(stageIndex As Integer)
    Dim instruction As String
    instruction = Pipeline(stageIndex).instruction
    instruction = Split(instruction, ";")(0)
    instruction = Trim(instruction)
    
    Dim parts() As String
    parts = Split(instruction, " ")
    
    If UBound(parts) >= 0 Then
        Pipeline(stageIndex).Opcode = UCase(Replace(Trim(parts(0)), ",", ""))
    End If
    
    Dim operands As String
    If UBound(parts) >= 1 Then
        operands = Trim(parts(1))
        If UBound(parts) >= 2 Then
            operands = operands & " " & Trim(parts(2))
        End If
        operands = Replace(operands, ",", "")
        Dim opParts() As String
        opParts = Split(Trim(operands), " ")
        
        If UBound(opParts) >= 0 Then Pipeline(stageIndex).Operand1 = Trim(opParts(0))
        If UBound(opParts) >= 1 Then Pipeline(stageIndex).Operand2 = Trim(opParts(1))
    End If
End Sub

Sub ExecuteInstruction(stageIndex As Integer)
    Dim op As String
    op = Pipeline(stageIndex).Opcode
    
    Select Case op
        Case "MOV", "LOAD"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " ? " & Pipeline(stageIndex).Operand2
        Case "ADD"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " + " & Pipeline(stageIndex).Operand2
        Case "SUB"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " - " & Pipeline(stageIndex).Operand2
        Case "MUL"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " × " & Pipeline(stageIndex).Operand2
        Case "DIV"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " ÷ " & Pipeline(stageIndex).Operand2
        Case "NOP"
            Pipeline(stageIndex).Result = "No Operation"
        Case Else
            Pipeline(stageIndex).Result = "Execute: " & op
    End Select
End Sub

Function AllInstructionsCompleted() As Boolean
    ' Verificar si hay más instrucciones por cargar
    If CurrentInstructionIndex <= UBound(Instructions) Then
        AllInstructionsCompleted = False
        Exit Function
    End If
    
    ' Verificar si hay instrucciones en el pipeline
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            AllInstructionsCompleted = False
            Exit Function
        End If
    Next i
    
    AllInstructionsCompleted = True
End Function

' =============================================
' VISUALIZACIÓN PIPELINE
' =============================================

Sub CreatePipelineDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Pipeline")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Pipeline"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Tab.Color = RGB(100, 150, 200)
    
    ' Encabezado principal
    With ws.Range("A1:H1")
        .Merge
        .value = "SIMULADOR DE PIPELINE - 5 ETAPAS"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(50, 80, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' Información del ciclo
    ws.Range("A3").value = "Ciclo de Reloj:"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").value = 0
    ws.Range("B3").Font.Size = 14
    ws.Range("B3").Font.Bold = True
    ws.Range("B3").Font.Color = RGB(200, 0, 0)
    
    ws.Range("D3").value = "Instrucciones:"
    ws.Range("D3").Font.Bold = True
    ws.Range("E3").value = GetInstructionCount()
    ws.Range("E3").Font.Size = 12
    ws.Range("E3").Font.Bold = True
    
    ' Encabezados de etapas
    Dim stages As Variant
    stages = Array("ETAPA", "IF", "ID", "EX", "MEM", "WB")
    Dim col As Integer
    For col = 0 To 5
        With ws.Cells(5, col + 1)
            .value = stages(col)
            .Font.Bold = True
            .Font.Size = 11
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(70, 100, 150)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.Weight = xlMedium
        End With
    Next col
    
    ' Descripciones de etapas
    With ws.Cells(6, 2)
        .value = "Instruction Fetch"
        .Font.Italic = True
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(6, 3)
        .value = "Instruction Decode"
        .Font.Italic = True
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(6, 4)
        .value = "Execute"
        .Font.Italic = True
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(6, 5)
        .value = "Memory Access"
        .Font.Italic = True
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(6, 6)
        .value = "Write Back"
        .Font.Italic = True
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    
    ' Filas para instrucciones
    Dim row As Integer
    For row = 7 To 11
        ws.ROWS(row).RowHeight = 35
        For col = 1 To 6
            With ws.Cells(row, col)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
        Next col
    Next row
    
    ' Ajustar columnas
    ws.Columns("A:A").ColumnWidth = 12
    ws.Columns("B:F").ColumnWidth = 15
End Sub

Sub UpdatePipelineDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pipeline")
    
    ' Actualizar ciclo
    ws.Range("B3").value = ClockCycle
    ws.Range("E3").value = GetInstructionCount()
    
    ' Limpiar área de pipeline
    Dim row As Integer, col As Integer
    For row = 7 To 11
        For col = 1 To 6
            ws.Cells(row, col).value = ""
            ws.Cells(row, col).Interior.Color = RGB(240, 240, 240)
            ws.Cells(row, col).Font.Bold = False
            ws.Cells(row, col).Font.Color = RGB(0, 0, 0)
        Next col
    Next row
    
    ' Mostrar instrucciones en pipeline
    Dim instructionRows As Object
    Set instructionRows = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            Dim instNum As Long
            instNum = Pipeline(i).InstructionNumber
            
            If Not instructionRows.Exists(instNum) Then
                instructionRows.Add instNum, instructionRows.count + 7
            End If
            
            Dim currentRow As Long
            currentRow = instructionRows(instNum)
            
            ' Mostrar en etapa correspondiente
            For col = 2 To 6
                Dim stageCol As String
                stageCol = ws.Cells(5, col).value
                
                If stageCol = Pipeline(i).stage Then
                    ' Etapa actual
                    ws.Cells(currentRow, col).value = "I" & instNum
                    ws.Cells(currentRow, col).Font.Bold = True
                    
                    If Pipeline(i).Stalled Then
                        ws.Cells(currentRow, col).Interior.Color = RGB(255, 100, 100)
                        ws.Cells(currentRow, col).Font.Color = RGB(255, 255, 255)
                    Else
                        ws.Cells(currentRow, col).Interior.Color = Pipeline(i).Color
                    End If
                ElseIf GetStageOrder(stageCol) < GetStageOrder(Pipeline(i).stage) Then
                    ' Etapa completada
                    ws.Cells(currentRow, col).value = "?"
                    ws.Cells(currentRow, col).Interior.Color = RGB(200, 255, 200)
                    ws.Cells(currentRow, col).Font.Size = 12
                    ws.Cells(currentRow, col).Font.Color = RGB(0, 100, 0)
                End If
            Next col
        End If
    Next i
End Sub

Function GetStageOrder(stage As String) As Integer
    Select Case stage
        Case "IF": GetStageOrder = 0
        Case "ID": GetStageOrder = 1
        Case "EX": GetStageOrder = 2
        Case "MEM": GetStageOrder = 3
        Case "WB": GetStageOrder = 4
        Case Else: GetStageOrder = -1
    End Select
End Function

Function GetInstructionColor(instNum As Long) As Long
    ' Colores predefinidos para instrucciones
    Select Case (instNum - 1) Mod 10
        Case 0: GetInstructionColor = RGB(173, 216, 230)  ' Azul claro
        Case 1: GetInstructionColor = RGB(255, 182, 193)  ' Rosa claro
        Case 2: GetInstructionColor = RGB(221, 160, 221)  ' Violeta
        Case 3: GetInstructionColor = RGB(255, 218, 185)  ' Durazno
        Case 4: GetInstructionColor = RGB(176, 224, 230)  ' Azul polvo
        Case 5: GetInstructionColor = RGB(240, 230, 140)  ' Caqui
        Case 6: GetInstructionColor = RGB(152, 251, 152)  ' Verde claro
        Case 7: GetInstructionColor = RGB(255, 228, 196)  ' Bisque
        Case 8: GetInstructionColor = RGB(230, 230, 250)  ' Lavanda
        Case 9: GetInstructionColor = RGB(245, 222, 179)  ' Trigo
    End Select
End Function

Sub ReiniciarPipeline()
    InitializePipeline
    CreatePipelineDisplay
    UpdatePipelineDisplay
End Sub

Sub PausarPipeline()
    IsPipelineRunning = False
End Sub
