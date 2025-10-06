Attribute VB_Name = "PipelineRiesgos"
' =============================================
' MÓDULO: PipelineRiesgos
' Descripción: Extensión del Pipeline para detección de riesgos de datos
' =============================================

Option Explicit

Type PipelineInstrRiesgos
    instruction As String
    Opcode As String
    Operand1 As String
    Operand2 As String
    Operand3 As String
    stage As String
    CycleEntered As Long
    CurrentStageCycle As Long
    Color As Long
    Result As String
    Stalled As Boolean
    StallCycles As Long
    DestinationReg As String
    SourceReg1 As String
    SourceReg2 As String
    InstructionNumber As Long
End Type

' Variables globales únicas para este módulo
Dim PipelineR() As PipelineInstrRiesgos
Dim ClockCycleR As Long
Dim InstructionsR() As String
Dim CurrentInstructionIndexR As Long
Dim PipelineStagesR(4) As String
Dim IsPipelineRunningR As Boolean
Dim RegisterStatusR(15) As String
Dim TotalStallCyclesR As Long

' =============================================
' INICIALIZACIÓN DEL PIPELINE CON RIESGOS
' =============================================

Public Sub IniciarSimuladorConRiesgos()
    InitializePipelineRiesgos
    MsgBox "? Simulador de Pipeline con Detección de Riesgos inicializado." & vbCrLf & _
           "Se detectarán automáticamente riesgos RAW y se insertarán burbujas.", vbInformation
End Sub

Public Sub ReiniciarPipelineRiesgos()
    InitializePipelineRiesgos
    MsgBox "?? Pipeline con riesgos reiniciado correctamente.", vbInformation
End Sub

Private Sub InitializePipelineRiesgos()
    ' Configurar etapas del pipeline
    PipelineStagesR(0) = "IF"
    PipelineStagesR(1) = "ID"
    PipelineStagesR(2) = "EX"
    PipelineStagesR(3) = "MEM"
    PipelineStagesR(4) = "WB"
    
    ' Inicializar pipeline vacío
    ReDim PipelineR(0 To 4)
    Dim i As Integer
    For i = 0 To 4
        ClearPipelineStageR i
    Next i
    
    ' Inicializar estado de registros
    For i = 0 To 15
        RegisterStatusR(i) = "READY"
    Next i
    
    ClockCycleR = 0
    CurrentInstructionIndexR = 0
    TotalStallCyclesR = 0
    
    ' Cargar instrucciones desde hoja
    LoadInstructionsFromSheetR
    
    CreateUnifiedPipelineDisplayR
    UpdatePipelineDisplayR
    
    LogMessageR "?? Pipeline con riesgos inicializado con " & GetInstructionCountR() & " instrucciones"
End Sub

Private Sub ClearPipelineStageR(stageIndex As Integer)
    PipelineR(stageIndex).instruction = ""
    PipelineR(stageIndex).Opcode = ""
    PipelineR(stageIndex).Operand1 = ""
    PipelineR(stageIndex).Operand2 = ""
    PipelineR(stageIndex).Operand3 = ""
    PipelineR(stageIndex).stage = ""
    PipelineR(stageIndex).CycleEntered = 0
    PipelineR(stageIndex).CurrentStageCycle = 0
    PipelineR(stageIndex).Color = RGB(255, 255, 255)
    PipelineR(stageIndex).Result = ""
    PipelineR(stageIndex).Stalled = False
    PipelineR(stageIndex).StallCycles = 0
    PipelineR(stageIndex).DestinationReg = ""
    PipelineR(stageIndex).SourceReg1 = ""
    PipelineR(stageIndex).SourceReg2 = ""
    PipelineR(stageIndex).InstructionNumber = 0
End Sub

Private Sub LoadInstructionsFromSheetR()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PipelineCodigo")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Instrucciones de ejemplo que causan riesgos
        ReDim InstructionsR(0 To 5)
        InstructionsR(0) = "ADD R1, R2, R3"
        InstructionsR(1) = "SUB R4, R1, R5"    ' RAW hazard con R1
        InstructionsR(2) = "MUL R6, R7, R8"
        InstructionsR(3) = "DIV R9, R1, R10"   ' RAW hazard con R1
        InstructionsR(4) = "MOV R11, R12"
        InstructionsR(5) = "ADD R13, R4, R1"   ' RAW hazard con R1 y R4
    Else
        ' Cargar instrucciones desde la hoja
        Dim lastRow As Long
        lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
        If lastRow < 2 Then
            ' Si no hay instrucciones, usar las de ejemplo
            ReDim InstructionsR(0 To 5)
            InstructionsR(0) = "ADD R1, R2, R3"
            InstructionsR(1) = "SUB R4, R1, R5"
            InstructionsR(2) = "MUL R6, R7, R8"
            InstructionsR(3) = "DIV R9, R1, R10"
            InstructionsR(4) = "MOV R11, R12"
            InstructionsR(5) = "ADD R13, R4, R1"
        Else
            ReDim InstructionsR(0 To lastRow - 2)
            Dim i As Long
            For i = 2 To lastRow
                If i - 2 <= UBound(InstructionsR) Then
                    InstructionsR(i - 2) = Trim(ws.Cells(i, 1).value)
                End If
            Next i
        End If
    End If
    
    ' Verificar que el array no esté vacío
    If GetInstructionCountR() = 0 Then
        ReDim InstructionsR(0 To 0)
        InstructionsR(0) = "NOP"
    End If
    
    LogMessageR "?? Cargadas " & GetInstructionCountR() & " instrucciones para pipeline con riesgos"
End Sub
' CORRECCIÓN: Función GetInstructionCountR mejorada
Private Function GetInstructionCountR() As Long
    On Error GoTo ErrorHandler
    
    ' Verificar si el array está inicializado correctamente
    If Not IsArrayInitialized(InstructionsR) Then
        GetInstructionCountR = 0
        Exit Function
    End If
    
    GetInstructionCountR = UBound(InstructionsR) - LBound(InstructionsR) + 1
    Exit Function
    
ErrorHandler:
    GetInstructionCountR = 0
End Function

' Función auxiliar para verificar si un array está inicializado
Private Function IsArrayInitialized(arr As Variant) As Boolean
    On Error GoTo ErrorHandler
    ' Intentar acceder al límite del array
    Dim test As Long
    test = UBound(arr)
    IsArrayInitialized = True
    Exit Function
    
ErrorHandler:
    IsArrayInitialized = False
End Function
' =============================================
' DETECCIÓN DE RIESGOS DE DATOS - CORREGIDA
' =============================================

Private Function CheckForDataHazardsR(currentStageIndex As Integer) As String
    If currentStageIndex <> 1 Then ' Solo verificar en etapa ID
        CheckForDataHazardsR = ""
        Exit Function
    End If
    
    ' Si no hay instrucción en ID, salir
    If PipelineR(currentStageIndex).instruction = "" Then
        CheckForDataHazardsR = ""
        Exit Function
    End If
    
    Dim currentInstr As PipelineInstrRiesgos
    currentInstr = PipelineR(currentStageIndex)
    Dim hazardType As String
    
    ' Verificar dependencias con instrucción en EX
    If PipelineR(2).stage = "EX" And PipelineR(2).instruction <> "" Then
        hazardType = CheckRAWHazardR(currentInstr, PipelineR(2), "EX")
        If hazardType <> "" Then
            CheckForDataHazardsR = "RAW-EX: " & hazardType
            Exit Function
        End If
    End If
    
    ' Verificar dependencias con instrucción en MEM
    If PipelineR(3).stage = "MEM" And PipelineR(3).instruction <> "" Then
        hazardType = CheckRAWHazardR(currentInstr, PipelineR(3), "MEM")
        If hazardType <> "" Then
            CheckForDataHazardsR = "RAW-MEM: " & hazardType
            Exit Function
        End If
    End If
    
    CheckForDataHazardsR = ""
End Function

Private Function CheckRAWHazardR(currentInstr As PipelineInstrRiesgos, previousInstr As PipelineInstrRiesgos, stage As String) As String
    Dim dependency As String
    
    If previousInstr.DestinationReg <> "" Then
        ' Verificar dependencia en primer operando fuente
        If currentInstr.SourceReg1 <> "" And currentInstr.SourceReg1 = previousInstr.DestinationReg Then
            dependency = currentInstr.SourceReg1 & " (Operando 1)"
        End If
        
        ' Verificar dependencia en segundo operando fuente
        If currentInstr.SourceReg2 <> "" And currentInstr.SourceReg2 = previousInstr.DestinationReg Then
            If dependency <> "" Then dependency = dependency & ", "
            dependency = dependency & currentInstr.SourceReg2 & " (Operando 2)"
        End If
    End If
    
    CheckRAWHazardR = dependency
End Function

Private Sub HandleDataHazardR(stageIndex As Integer, hazardMessage As String)
    PipelineR(stageIndex).Stalled = True
    PipelineR(stageIndex).StallCycles = PipelineR(stageIndex).StallCycles + 1
    TotalStallCyclesR = TotalStallCyclesR + 1
    
    ' También detener la etapa IF si ID está detenida
    If stageIndex = 1 Then
        PipelineR(0).Stalled = True
    End If
    
    LogHazardR hazardMessage, PipelineR(stageIndex).instruction, GetAffectedInstructionR(hazardMessage)
    UpdateHazardDisplayR hazardMessage
End Sub

Private Function GetAffectedInstructionR(hazardMessage As String) As String
    If InStr(hazardMessage, "EX") > 0 And PipelineR(2).stage = "EX" Then
        GetAffectedInstructionR = PipelineR(2).instruction
    ElseIf InStr(hazardMessage, "MEM") > 0 And PipelineR(3).stage = "MEM" Then
        GetAffectedInstructionR = PipelineR(3).instruction
    Else
        GetAffectedInstructionR = "Desconocida"
    End If
End Function

Private Sub ResolveHazardsR()
    Dim i As Integer
    For i = 0 To 4
        If PipelineR(i).Stalled Then
            PipelineR(i).StallCycles = PipelineR(i).StallCycles + 1
            PipelineR(i).Result = "BURBUJA (Ciclo " & PipelineR(i).StallCycles & ")"
            
            ' Resolver después de 1 ciclo de burbuja
            If PipelineR(i).StallCycles >= 1 Then
                PipelineR(i).Stalled = False
                PipelineR(i).StallCycles = 0
                PipelineR(i).Result = "Burbuja resuelta"
                LogMessageR "?? Burbuja resuelta para: " & PipelineR(i).instruction
            End If
        End If
    Next i
End Sub

' =============================================
' SIMULACIÓN DEL PIPELINE CON DETECCIÓN DE HAZARDS - CORREGIDA
' =============================================

Public Sub EjecutarPipelineRiesgosCompleto()
    IsPipelineRunningR = True
    Dim maxCycles As Integer
    maxCycles = 30
    
    Do While IsPipelineRunningR And ClockCycleR < maxCycles
        ClockCycleR = ClockCycleR + 1
        AdvancePipelineR
        UpdatePipelineDisplayR
        DoEvents
        
        ' REEMPLAZAR Application.Wait con una alternativa más confiable
        PausaPersonalizada 0.5 ' 500 milisegundos
        
        If AllInstructionsCompletedR() Then
            IsPipelineRunningR = False
            LogMessageR "?? SIMULACIÓN COMPLETADA - Todas las instrucciones finalizadas"
            MsgBox "Simulación completada en " & ClockCycleR & " ciclos" & vbCrLf & _
                   "Ciclos de stall: " & TotalStallCyclesR & vbCrLf & _
                   "Eficiencia: " & Format((ClockCycleR - TotalStallCyclesR) / ClockCycleR, "0.0%"), vbInformation
        End If
    Loop
    
    If ClockCycleR >= maxCycles Then
        MsgBox "Límite de ciclos alcanzado", vbInformation
    End If
    
    IsPipelineRunningR = False
End Sub

Private Sub PausaPersonalizada(segundos As Double)
    Dim tiempoInicio As Double
    tiempoInicio = Timer
    Do While Timer < tiempoInicio + segundos
        DoEvents
    Loop
End Sub

Public Sub AvanzarCicloRiesgos()
    If IsPipelineRunningR Then
        MsgBox "Detenga la simulación automática primero", vbExclamation
        Exit Sub
    End If
    
    ClockCycleR = ClockCycleR + 1
    AdvancePipelineR
    UpdatePipelineDisplayR
    
    If AllInstructionsCompletedR() Then
        MsgBox "? Todas las instrucciones completadas en ciclo " & ClockCycleR, vbInformation
    End If
End Sub

' CORRECCIÓN: Lógica de avance del pipeline mejorada
Private Sub AdvancePipelineR()
    ' Primero resolver hazards existentes
    ResolveHazardsR
    
    ' Avanzar etapas de atrás hacia adelante
    AdvanceStageR "WB"
    AdvanceStageR "MEM"
    AdvanceStageR "EX"
    AdvanceStageR "ID"
    AdvanceStageR "IF"
    
    ' Insertar nueva instrucción si es posible
    If CurrentInstructionIndexR < GetInstructionCountR() Then
        If PipelineR(0).stage = "" And Not PipelineR(0).Stalled Then
            InsertNewInstructionR InstructionsR(CurrentInstructionIndexR)
            CurrentInstructionIndexR = CurrentInstructionIndexR + 1
        End If
    End If
    
    ' Verificar hazards después del avance
    CheckHazardsAfterAdvanceR
End Sub

Private Sub CheckHazardsAfterAdvanceR()
    Dim i As Integer
    For i = 0 To 4
        If PipelineR(i).stage = "ID" And Not PipelineR(i).Stalled And PipelineR(i).instruction <> "" Then
            Dim hazardMessage As String
            hazardMessage = CheckForDataHazardsR(i)
            If hazardMessage <> "" Then
                HandleDataHazardR i, hazardMessage
            End If
        End If
    Next i
End Sub

Private Function HasActiveStallsR() As Boolean
    Dim i As Integer
    For i = 0 To 4
        If PipelineR(i).Stalled Then
            HasActiveStallsR = True
            Exit Function
        End If
    Next i
    HasActiveStallsR = False
End Function

Private Sub AdvanceStageR(stage As String)
    Dim stageIndex As Integer
    stageIndex = GetStageIndexR(stage)
    
    If stageIndex = -1 Then Exit Sub
    
    ' Solo avanzar si la etapa está ocupada y no está detenida
    If PipelineR(stageIndex).stage = stage And Not PipelineR(stageIndex).Stalled And PipelineR(stageIndex).instruction <> "" Then
        PipelineR(stageIndex).CurrentStageCycle = PipelineR(stageIndex).CurrentStageCycle + 1
        
        If CanAdvanceToNextStageR(stageIndex) Then
            MoveToNextStageR stageIndex
        End If
    End If
End Sub

Private Function CanAdvanceToNextStageR(currentStageIndex As Integer) As Boolean
    Dim nextStageIndex As Integer
    nextStageIndex = currentStageIndex + 1
    
    If nextStageIndex > 4 Then
        CanAdvanceToNextStageR = True ' WB puede completarse
        Exit Function
    End If
    
    ' Verificar si la siguiente etapa está libre
    If PipelineR(nextStageIndex).stage = "" And Not PipelineR(nextStageIndex).Stalled Then
        CanAdvanceToNextStageR = True
    Else
        CanAdvanceToNextStageR = False
    End If
End Function

Private Sub MoveToNextStageR(currentStageIndex As Integer)
    Dim nextStageIndex As Integer
    nextStageIndex = currentStageIndex + 1
    
    If nextStageIndex <= 4 Then
        ' Mover a la siguiente etapa
        PipelineR(nextStageIndex) = PipelineR(currentStageIndex)
        PipelineR(nextStageIndex).stage = PipelineStagesR(nextStageIndex)
        PipelineR(nextStageIndex).CurrentStageCycle = 0
        
        ProcessInstructionInStageR nextStageIndex
    Else
        ' Completar la instrucción
        PipelineR(currentStageIndex).stage = "DONE"
        LogMessageR "? Instrucción completada: " & PipelineR(currentStageIndex).instruction
    End If
    
    ClearPipelineStageR currentStageIndex
End Sub

Private Sub InsertNewInstructionR(instruction As String)
    PipelineR(0).instruction = instruction
    PipelineR(0).stage = "IF"
    PipelineR(0).CycleEntered = ClockCycleR
    PipelineR(0).CurrentStageCycle = 0
    PipelineR(0).Color = GetInstructionColorR(CurrentInstructionIndexR + 1)
    PipelineR(0).Stalled = False
    PipelineR(0).InstructionNumber = CurrentInstructionIndexR + 1
    
    ProcessInstructionInStageR 0
    LogMessageR "?? Nueva instrucción: " & instruction
End Sub

Private Sub ProcessInstructionInStageR(stageIndex As Integer)
    Select Case PipelineStagesR(stageIndex)
        Case "IF"
            PipelineR(stageIndex).Result = "Capturando instrucción"
            
        Case "ID"
            ParseInstructionR stageIndex
            PipelineR(stageIndex).Result = "Decodificando: " & PipelineR(stageIndex).Opcode
            
        Case "EX"
            ExecuteInstructionR stageIndex
            PipelineR(stageIndex).Result = "Ejecutando: " & PipelineR(stageIndex).Result
            
        Case "MEM"
            PipelineR(stageIndex).Result = "Acceso a memoria"
            
        Case "WB"
            PipelineR(stageIndex).Result = "Escritura de resultado"
            ' Liberar el registro de destino
            If PipelineR(stageIndex).DestinationReg <> "" Then
                RegisterStatusR(GetRegisterIndex(PipelineR(stageIndex).DestinationReg)) = "READY"
            End If
    End Select
End Sub

' Función auxiliar para obtener índice del registro
Private Function GetRegisterIndex(reg As String) As Integer
    If Len(reg) > 1 And UCase(left(reg, 1)) = "R" Then
        On Error GoTo ErrorHandler
        GetRegisterIndex = CInt(Mid(reg, 2))
        Exit Function
    End If
ErrorHandler:
    GetRegisterIndex = 0
End Function

Private Sub ParseInstructionR(stageIndex As Integer)
    Dim instruction As String
    instruction = PipelineR(stageIndex).instruction
    
    If InStr(instruction, ";") > 0 Then
        instruction = Trim(left(instruction, InStr(instruction, ";") - 1))
    End If
    
    Dim parts() As String
    parts = Split(Trim(instruction), " ")
    
    If UBound(parts) >= 0 Then
        PipelineR(stageIndex).Opcode = UCase(Trim(parts(0)))
    End If
    
    ExtractOperandsR stageIndex, parts
End Sub

Private Sub ExtractOperandsR(stageIndex As Integer, parts() As String)
    PipelineR(stageIndex).DestinationReg = ""
    PipelineR(stageIndex).SourceReg1 = ""
    PipelineR(stageIndex).SourceReg2 = ""
    
    Select Case PipelineR(stageIndex).Opcode
        Case "MOV", "LOAD"
            If UBound(parts) >= 2 Then
                PipelineR(stageIndex).DestinationReg = ExtractRegisterR(parts(1))
                PipelineR(stageIndex).Operand1 = parts(2)
            End If
            
        Case "ADD", "SUB", "MUL", "DIV", "AND", "OR"
            If UBound(parts) >= 3 Then
                PipelineR(stageIndex).DestinationReg = ExtractRegisterR(parts(1))
                PipelineR(stageIndex).SourceReg1 = ExtractRegisterR(parts(2))
                PipelineR(stageIndex).SourceReg2 = ExtractRegisterR(parts(3))
            End If
            
        Case "STORE"
            If UBound(parts) >= 2 Then
                PipelineR(stageIndex).SourceReg1 = ExtractRegisterR(parts(1))
                PipelineR(stageIndex).Operand1 = parts(2)
            End If
    End Select
End Sub

Private Function ExtractRegisterR(operand As String) As String
    operand = Replace(operand, ",", "")
    If UCase(left(operand, 1)) = "R" Then
        ExtractRegisterR = UCase(Trim(operand))
    Else
        ExtractRegisterR = ""
    End If
End Function

Private Sub ExecuteInstructionR(stageIndex As Integer)
    Dim op As String
    op = PipelineR(stageIndex).Opcode
    
    Select Case op
        Case "MOV", "LOAD"
            PipelineR(stageIndex).Result = PipelineR(stageIndex).DestinationReg & " ? " & PipelineR(stageIndex).Operand1
        Case "ADD"
            PipelineR(stageIndex).Result = PipelineR(stageIndex).SourceReg1 & " + " & PipelineR(stageIndex).SourceReg2
        Case "SUB"
            PipelineR(stageIndex).Result = PipelineR(stageIndex).SourceReg1 & " - " & PipelineR(stageIndex).SourceReg2
        Case "MUL"
            PipelineR(stageIndex).Result = PipelineR(stageIndex).SourceReg1 & " × " & PipelineR(stageIndex).SourceReg2
        Case "DIV"
            PipelineR(stageIndex).Result = PipelineR(stageIndex).SourceReg1 & " ÷ " & PipelineR(stageIndex).SourceReg2
        Case Else
            PipelineR(stageIndex).Result = "Operación: " & op
    End Select
End Sub

Private Function AllInstructionsCompletedR() As Boolean
    ' Verificar si hay más instrucciones por cargar
    If CurrentInstructionIndexR < GetInstructionCountR() Then
        AllInstructionsCompletedR = False
        Exit Function
    End If
    
    ' Verificar si el pipeline está vacío
    Dim i As Integer
    For i = 0 To 4
        If PipelineR(i).stage <> "" And PipelineR(i).stage <> "DONE" Then
            AllInstructionsCompletedR = False
            Exit Function
        End If
    Next i
    
    AllInstructionsCompletedR = True
End Function

Private Function GetStageIndexR(stage As String) As Integer
    Dim i As Integer
    For i = 0 To 4
        If PipelineStagesR(i) = stage Then
            GetStageIndexR = i
            Exit Function
        End If
    Next i
    GetStageIndexR = -1
End Function

' =============================================
' INTERFAZ UNIFICADA CON VISUALIZACIÓN DE RIESGOS
' =============================================
Private Sub CreateUnifiedPipelineDisplayR()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PipelineRiesgos")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "PipelineRiesgos"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Tab.Color = RGB(70, 130, 180)
    
    With ws.Range("A1:H1")
        .Merge
        .value = "?? PIPELINE CON DETECCIÓN DE RIESGOS DE DATOS"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(70, 130, 180)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    ws.Range("A3").value = "Ciclo Actual:"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").value = ClockCycleR
    ws.Range("B3").Font.Size = 14
    ws.Range("B3").Font.Bold = True
    
    ws.Range("D3").value = "Instrucciones:"
    ws.Range("D3").Font.Bold = True
    ws.Range("E3").value = GetInstructionCountR()
    
    ws.Range("G3").value = "Burbujas:"
    ws.Range("G3").Font.Bold = True
    ws.Range("H3").value = TotalStallCyclesR
    ws.Range("H3").Interior.Color = RGB(255, 200, 200)
    
    ws.Range("A5").value = "?? DIAGRAMA DEL PIPELINE"
    ws.Range("A5").Font.Bold = True
    ws.Range("A5").Font.Size = 12
    
    Dim stages As Variant
    stages = Array("Inst#", "IF", "ID", "EX", "MEM", "WB", "Estado", "Riesgos")
    
    Dim col As Integer
    For col = 0 To 7
        With ws.Cells(7, col + 1)
            .value = stages(col)
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(70, 130, 180)
            .HorizontalAlignment = xlCenter
            .Borders.Weight = xlThin
        End With
    Next col
    
    Dim row As Integer
    For row = 8 To 12
        ws.ROWS(row).RowHeight = 35
        For col = 1 To 8
            With ws.Cells(row, col)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
        Next col
    Next row
    
    ws.Range("A14").value = "?? RIESGOS DETECTADOS"
    ws.Range("A14").Font.Bold = True
    ws.Range("A14").Font.Size = 12
    
    Dim hazardHeaders As Variant
    hazardHeaders = Array("Ciclo", "Tipo", "Instrucción Afectada", "Instrucción en Conflicto", "Registros", "Acción")
    
    For col = 0 To 5
        With ws.Cells(15, col + 1)
            .value = hazardHeaders(col)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(178, 34, 34)
            .HorizontalAlignment = xlCenter
        End With
    Next col
    
    ws.Range("A25").value = "?? LOG DE EVENTOS"
    ws.Range("A25").Font.Bold = True
    ws.Range("A25").Font.Size = 12
    
    CreatePipelineButtonsR ws
    
    ws.Columns("A:A").ColumnWidth = 8
    ws.Columns("B:F").ColumnWidth = 15
    ws.Columns("G:G").ColumnWidth = 20
    ws.Columns("H:H").ColumnWidth = 15
End Sub

Private Sub CreatePipelineButtonsR(ws As Worksheet)
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    Dim btn As Button
    Set btn = ws.Buttons.Add(50, 400, 100, 30)
    btn.OnAction = "EjecutarPipelineRiesgosCompleto"
    btn.Characters.text = "?? Ejecutar"
    
    Set btn = ws.Buttons.Add(160, 400, 100, 30)
    btn.OnAction = "AvanzarCicloRiesgos"
    btn.Characters.text = "?? Avanzar"
    
    Set btn = ws.Buttons.Add(270, 400, 80, 30)
    btn.OnAction = "ReiniciarPipelineRiesgos"
    btn.Characters.text = "?? Reiniciar"
End Sub

Private Sub UpdatePipelineDisplayR()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PipelineRiesgos")
    
    ws.Range("B3").value = ClockCycleR
    ws.Range("E3").value = GetInstructionCountR()
    ws.Range("H3").value = TotalStallCyclesR
    
    Dim row As Integer, col As Integer
    For row = 8 To 12
        For col = 1 To 8
            ws.Cells(row, col).value = ""
            ws.Cells(row, col).Interior.Color = RGB(240, 240, 240)
            ws.Cells(row, col).Font.Bold = False
        Next col
    Next row
    
    Dim i As Integer
    For i = 0 To 4
        If PipelineR(i).stage <> "" Then
            ws.Cells(8 + i, 1).value = "I" & PipelineR(i).InstructionNumber
            ws.Cells(8 + i, 1).Font.Bold = True
            
            Dim stageCol As Integer
            stageCol = GetStageColumnR(PipelineR(i).stage)
            
            If stageCol > 0 Then
                ws.Cells(8 + i, stageCol).value = PipelineR(i).instruction
                ws.Cells(8 + i, stageCol).Interior.Color = PipelineR(i).Color
                ws.Cells(8 + i, stageCol).Font.Bold = True
                
                ws.Cells(8 + i, 7).value = PipelineR(i).Result
                
                If PipelineR(i).Stalled Then
                    ws.Cells(8 + i, 8).value = "?? BURBUJA ACTIVA"
                    ws.Cells(8 + i, 8).Interior.Color = RGB(255, 100, 100)
                    ws.Cells(8 + i, 8).Font.Bold = True
                    For col = 1 To 8
                        ws.Cells(8 + i, col).Interior.Color = RGB(255, 200, 200)
                    Next col
                Else
                    ws.Cells(8 + i, 8).value = "? Sin riesgos"
                    ws.Cells(8 + i, 8).Interior.Color = RGB(200, 255, 200)
                End If
            End If
        End If
    Next i
End Sub

Private Function GetStageColumnR(stage As String) As Integer
    Select Case stage
        Case "IF": GetStageColumnR = 2
        Case "ID": GetStageColumnR = 3
        Case "EX": GetStageColumnR = 4
        Case "MEM": GetStageColumnR = 5
        Case "WB": GetStageColumnR = 6
        Case Else: GetStageColumnR = 0
    End Select
End Function

Private Sub UpdateHazardDisplayR(hazardMessage As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PipelineRiesgos")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    If nextRow < 16 Then nextRow = 16
    
    ws.Cells(nextRow, 1).value = ClockCycleR
    ws.Cells(nextRow, 2).value = hazardMessage
    ws.Cells(nextRow, 3).value = PipelineR(1).instruction
    ws.Cells(nextRow, 4).value = GetAffectedInstructionR(hazardMessage)
    ws.Cells(nextRow, 5).value = ExtractRegistersFromHazardR(hazardMessage)
    ws.Cells(nextRow, 6).value = "INSERTAR BURBUJA"
    
    ws.ROWS(nextRow).Interior.Color = RGB(255, 200, 200)
    ws.ROWS(nextRow).Font.Bold = True
End Sub

Private Function ExtractRegistersFromHazardR(hazardMessage As String) As String
    If InStr(hazardMessage, "(") > 0 Then
        Dim regPart As String
        regPart = Mid(hazardMessage, InStr(hazardMessage, "(") + 1)
        regPart = left(regPart, InStr(regPart, ")") - 1)
        ExtractRegistersFromHazardR = regPart
    Else
        ExtractRegistersFromHazardR = "N/A"
    End If
End Function

Private Function GetInstructionColorR(instNum As Long) As Long
    Dim colors As Variant
    colors = Array( _
        RGB(173, 216, 230), _
        RGB(255, 182, 193), _
        RGB(221, 160, 221), _
        RGB(255, 218, 185), _
        RGB(176, 224, 230), _
        RGB(240, 230, 140), _
        RGB(152, 251, 152), _
        RGB(255, 228, 196) _
    )
    GetInstructionColorR = colors((instNum - 1) Mod 8)
End Function

' =============================================
' LOGGING
' =============================================

Private Sub LogMessageR(message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PipelineRiesgos")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    If nextRow < 26 Then nextRow = 26
    
    ws.Cells(nextRow, 1).value = ClockCycleR
    ws.Cells(nextRow, 2).value = message
    
    If InStr(message, "?") > 0 Or InStr(message, "??") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(200, 255, 200)
    ElseIf InStr(message, "??") > 0 Or InStr(message, "??") > 0 Or InStr(message, "BURBUJA") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(255, 200, 200)
    ElseIf InStr(message, "??") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(200, 220, 255)
    Else
        ws.ROWS(nextRow).Interior.Color = RGB(240, 240, 240)
    End If
End Sub

Private Sub LogHazardR(hazardMessage As String, currentInstr As String, conflictInstr As String)
    LogMessageR "?? RIESGO: " & hazardMessage & " | " & currentInstr & " espera por " & conflictInstr
End Sub

' =============================================
' FUNCIONES PÚBLICAS ADICIONALES
' =============================================

Public Sub CargarProgramaPipelineRiesgos()
    LoadInstructionsFromSheetR
    UpdatePipelineDisplayR
    MsgBox "?? Programa recargado correctamente. " & GetInstructionCountR() & " instrucciones listas.", vbInformation
End Sub

Public Sub DetenerEjecucionRiesgos()
    IsPipelineRunningR = False
    LogMessageR "?? Ejecución detenida por el usuario"
    MsgBox "Ejecución detenida", vbInformation
End Sub

