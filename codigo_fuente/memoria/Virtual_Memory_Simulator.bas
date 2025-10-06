Attribute VB_Name = "Virtual_Memory_Simulator"
' M�dulo: Virtual_Memory_Simulator
' Descripci�n: Simulador de memoria virtual completo y funcional
' Versi�n: 4.1 (C�digo corregido y optimizado)

Option Explicit

' =================== CONSTANTES ===================
Private Const VM_SIZE As Integer = 256
Private Const VM_ROWS As Integer = 16
Private Const VM_COLS As Integer = 16

' =================== TIPOS ===================
Type MemoryCell
    address As String
    value As String
    instruction As String
    dataType As String
    Accessed As Boolean
    Modified As Boolean
End Type

' =================== VARIABLES GLOBALES ===================
Private VirtualMemory() As MemoryCell
Private CurrentInstructionVM As Integer
Private StackPointerVM As Integer
Private AX As Long, BX As Long, CX As Long, DX As Long
Private SI As Long, DI As Long, BP As Long, SP As Long
Private flags As Long
Private IsRunning As Boolean

' ===================================================================================
' INICIALIZACI�N
' ===================================================================================

Sub InitializeVirtualMemorySimulator()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ReDim VirtualMemory(0 To VM_SIZE - 1)
    
    CurrentInstructionVM = 0
    StackPointerVM = VM_SIZE - 1
    AX = 0: BX = 0: CX = 0: DX = 0
    SI = 0: DI = 0: BP = 0: SP = VM_SIZE - 1
    flags = 0
    IsRunning = False
    
    Dim i As Integer
    For i = 0 To VM_SIZE - 1
        VirtualMemory(i).address = "0x" & Format(Hex(i), "00")
        VirtualMemory(i).value = "00"
        VirtualMemory(i).instruction = ""
        VirtualMemory(i).dataType = "FREE"
        VirtualMemory(i).Accessed = False
        VirtualMemory(i).Modified = False
    Next i
    
    For i = VM_SIZE - 32 To VM_SIZE - 1
        VirtualMemory(i).dataType = "STACK"
    Next i
    
    CreateVMInterface
    CreateVMGestionSheet ' NUEVO: Crear hoja de gesti�n
    LoadProgramFromSheet
    UpdateVMDisplay
    UpdateRegisterDisplayVM
    UpdateGestionVMSheet ' NUEVO: Actualizar hoja de gesti�n
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Simulador de Memoria Virtual inicializado correctamente.", vbInformation
End Sub

Sub ResetVirtualMemory()
    CurrentInstructionVM = 0
    InitializeVirtualMemorySimulator
End Sub
Sub CreateVMGestionSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("GestionVM")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "GestionVM"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws
        ' T�tulo principal
        .Range("A1:F1").Merge
        .Range("A1").value = "GESTI�N DE MEMORIA VIRTUAL - TABLA COMPLETA"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 180, 220)
        
        ' Encabezados de columnas
        .Range("A2").value = "DIRECCI�N"
        .Range("B2").value = "VALOR"
        .Range("C2").value = "INSTRUCCI�N"
        .Range("D2").value = "TIPO"
        .Range("E2").value = "ACCEDIDO"
        .Range("F2").value = "MODIFICADO"
        
        ' Formato de encabezados
        With .Range("A2:F2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Configurar columnas
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 8
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 10
        .Columns("F").ColumnWidth = 10
        
        ' Estad�sticas
        .Range("H1").value = "ESTAD�STICAS DE MEMORIA"
        .Range("H1:K1").Merge
        .Range("H1").Font.Bold = True
        .Range("H1").Interior.Color = RGB(180, 220, 180)
        
        .Range("H2").value = "Total Celdas:"
        .Range("I2").value = VM_SIZE
        
        .Range("H3").value = "Instrucciones:"
        .Range("I3").value = "0"
        
        .Range("H4").value = "Datos:"
        .Range("I4").value = "0"
        
        .Range("H5").value = "Stack:"
        .Range("I5").value = "32"
        
        .Range("H6").value = "Libre:"
        .Range("I6").value = VM_SIZE - 32
    End With
End Sub

' Actualizar hoja de gesti�n con datos de la memoria virtual
' Actualizar hoja de gesti�n con datos de la memoria virtual
Sub UpdateGestionVMSheet()
    Dim ws As Worksheet
    Set ws = Worksheets("GestionVM")
    
    Dim i As Integer
    Dim instructionCount As Integer
    Dim dataCount As Integer
    Dim freeCount As Integer
    Dim stackCount As Integer
    
    instructionCount = 0
    dataCount = 0
    freeCount = 0
    stackCount = 0
    
    For i = 0 To VM_SIZE - 1
        ' Escribir datos en la tabla
        With VirtualMemory(i)
            ws.Cells(i + 3, 1).value = .address
            ws.Cells(i + 3, 2).value = .value
            ws.Cells(i + 3, 3).value = .instruction
            ws.Cells(i + 3, 4).value = .dataType
            ws.Cells(i + 3, 5).value = IIf(.Accessed, "?", "")
            ws.Cells(i + 3, 6).value = IIf(.Modified, "?", "")
        End With
        
        ' Aplicar formato condicional al rango
        With ws.Range(ws.Cells(i + 3, 1), ws.Cells(i + 3, 6))
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            
            ' Colores seg�n el tipo de memoria
            Select Case VirtualMemory(i).dataType
                Case "INSTR"
                    .Interior.Color = RGB(200, 255, 200) ' Verde para instrucciones
                    instructionCount = instructionCount + 1
                Case "DATA"
                    .Interior.Color = RGB(200, 220, 255) ' Azul para datos
                    dataCount = dataCount + 1
                Case "STACK"
                    .Interior.Color = RGB(255, 240, 200) ' Naranja para stack
                    stackCount = stackCount + 1
                Case "FREE"
                    .Interior.Color = RGB(240, 240, 240) ' Gris para libre
                    freeCount = freeCount + 1
            End Select
            
            ' Resaltar celdas accedidas
            If VirtualMemory(i).Accessed Then
                .Font.Bold = True
                .Font.Color = RGB(200, 0, 0)
            End If
            
            ' Resaltar celdas modificadas
            If VirtualMemory(i).Modified Then
                .Borders(xlEdgeBottom).Weight = xlThick
                .Borders(xlEdgeBottom).Color = RGB(255, 0, 0)
            End If
        End With
        
        ' Resaltar instrucci�n actual
        If i = CurrentInstructionVM And VirtualMemory(i).dataType = "INSTR" Then
            ws.Range(ws.Cells(i + 3, 1), ws.Cells(i + 3, 6)).Interior.Color = RGB(255, 255, 0) ' Amarillo
        End If
    Next i
    
    ' Actualizar estad�sticas
    With ws
        .Range("I3").value = instructionCount
        .Range("I4").value = dataCount
        .Range("I5").value = stackCount
        .Range("I6").value = freeCount
    End With
    
    ' Autoajustar filas
    ws.ROWS("3:" & VM_SIZE + 2).RowHeight = 20
End Sub ' ===================================================================================
' OPERACIONES DE MEMORIA
' ===================================================================================

Function WriteVMMemory(address As Integer, value As String, Optional dataType As String = "DATA") As Boolean
    On Error GoTo ErrorHandler
    If address < 0 Or address >= VM_SIZE Then
        WriteVMMemory = False
        Exit Function
    End If
    VirtualMemory(address).value = value
    VirtualMemory(address).dataType = dataType
    VirtualMemory(address).Modified = True
    VirtualMemory(address).Accessed = True
    WriteVMMemory = True
    
    ' NUEVO: Actualizar hoja de gesti�n
    UpdateGestionVMSheet
    Exit Function
ErrorHandler:
    WriteVMMemory = False
End Function

Function ReadVMMemory(address As Integer) As String
    On Error GoTo ErrorHandler
    If address < 0 Or address >= VM_SIZE Then
        ReadVMMemory = "00"
        Exit Function
    End If
    VirtualMemory(address).Accessed = True
    ReadVMMemory = VirtualMemory(address).value
    
    ' NUEVO: Actualizar hoja de gesti�n
    UpdateGestionVMSheet
    Exit Function
ErrorHandler:
    ReadVMMemory = "00"
End Function

' ===================================================================================
' CARGA DE PROGRAMAS
' ===================================================================================

Sub LoadProgramFromSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CodigoVM")
    On Error GoTo 0
    
    If ws Is Nothing Then
        CreateSampleVMProgram
        Set ws = ThisWorkbook.Sheets("CodigoVM")
    End If
    
    Dim i As Integer
    For i = 0 To VM_SIZE - 1
        If VirtualMemory(i).dataType = "INSTR" Then
            VirtualMemory(i).instruction = ""
            VirtualMemory(i).value = "00"
            VirtualMemory(i).dataType = "FREE"
        End If
    Next i
    
    Dim lastRow As Integer
    lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
    If lastRow < 2 Then lastRow = 2
    
    Dim instructionCount As Integer
    instructionCount = 0
    
    For i = 2 To lastRow
        Dim instruction As String
        instruction = Trim(ws.Cells(i, 1).value)
        
        If instruction <> "" And left(instruction, 1) <> ";" Then
            Dim addr As Integer
            addr = instructionCount
            
            If addr < VM_SIZE - 32 Then
                VirtualMemory(addr).instruction = instruction
                VirtualMemory(addr).value = left(instruction, 10)
                VirtualMemory(addr).dataType = "INSTR"
                instructionCount = instructionCount + 1
            End If
        End If
    Next i
    
    ' NUEVO: Actualizar hoja de gesti�n
    UpdateGestionVMSheet
    
    UpdateVMStatus "LISTO", "---", instructionCount & " instrucciones"
End Sub

Sub CreateSampleVMProgram()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "CodigoVM"
    
    ws.Range("A1").value = "Programa Ensamblador (una instrucci�n por l�nea)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    
    Dim row As Integer: row = 2
    ws.Cells(row, 1).value = "; Programa de ejemplo - Operaciones b�sicas": row = row + 1
    ws.Cells(row, 1).value = "MOV AX 10": row = row + 1
    ws.Cells(row, 1).value = "MOV BX 5": row = row + 1
    ws.Cells(row, 1).value = "ADD AX BX": row = row + 1
    ws.Cells(row, 1).value = "MOV CX 3": row = row + 1
    ws.Cells(row, 1).value = "MUL CX": row = row + 1
    ws.Cells(row, 1).value = "PUSH AX": row = row + 1
    ws.Cells(row, 1).value = "MOV DX 2": row = row + 1
    ws.Cells(row, 1).value = "POP BX": row = row + 1
    ws.Cells(row, 1).value = "SUB BX DX": row = row + 1
    ws.Cells(row, 1).value = "HLT": row = row + 1
    
    ws.Columns("A").ColumnWidth = 50
End Sub

' ===================================================================================
' EJECUCI�N
' ===================================================================================

Sub ExecuteNextInstructionVM()
    ' VALIDACI�N CR�TICA: Verificar l�mites antes de acceder al array
    If CurrentInstructionVM < 0 Or CurrentInstructionVM >= VM_SIZE Then
        UpdateVMStatus "ERROR", "---", "Puntero fuera de l�mites"
        MsgBox "Error: El puntero de instrucci�n est� fuera del rango v�lido.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Verificar si hay instrucci�n v�lida
    If VirtualMemory(CurrentInstructionVM).instruction = "" Or _
       VirtualMemory(CurrentInstructionVM).dataType <> "INSTR" Then
        UpdateVMStatus "COMPLETADO", "---", "No hay m�s instrucciones"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim instruction As String
    instruction = VirtualMemory(CurrentInstructionVM).instruction
    VirtualMemory(CurrentInstructionVM).Accessed = True
    
    UpdateVMStatus "EJECUTANDO", "0x" & Format(Hex(CurrentInstructionVM), "00"), instruction
    
    If ParseAndExecuteVM(instruction) Then
        UpdateVMDisplay
        HighlightCurrentInstructionVM
        UpdateRegisterDisplayVM
        UpdateGestionVMSheet ' NUEVO: Actualizar hoja de gesti�n
    Else
        UpdateVMStatus "ERROR", "0x" & Format(Hex(CurrentInstructionVM), "00"), "Error: " & instruction
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub ExecuteFullProgramVM()
    Application.ScreenUpdating = False
    IsRunning = True
    Dim stepCount As Integer: stepCount = 0
    Dim maxSteps As Integer: maxSteps = 50
    
    Do While IsRunning And stepCount < maxSteps
        ' Verificar l�mites del puntero de instrucci�n
        If CurrentInstructionVM < 0 Or CurrentInstructionVM >= VM_SIZE Then
            UpdateVMStatus "ERROR", "0x" & Hex(CurrentInstructionVM), "Puntero fuera de l�mites"
            IsRunning = False
            Exit Do
        End If
        
        ' Verificar si hay instrucci�n v�lida
        On Error Resume Next
        If VirtualMemory(CurrentInstructionVM).instruction = "" Or _
           VirtualMemory(CurrentInstructionVM).dataType <> "INSTR" Then
            IsRunning = False
            Exit Do
        End If
        If Err.Number <> 0 Then
            UpdateVMStatus "ERROR", "0x" & Hex(CurrentInstructionVM), "Error acceso memoria"
            IsRunning = False
            Exit Do
        End If
        On Error GoTo 0
        
        ' Ejecutar la instrucci�n
        ExecuteNextInstructionVM
        stepCount = stepCount + 1
        DoEvents
    Loop
    
    IsRunning = False
    Application.ScreenUpdating = True
    
    ' NUEVO: Actualizar hoja de gesti�n final
    UpdateGestionVMSheet
    
    If stepCount >= maxSteps Then
        MsgBox "Ejecuci�n detenida. L�mite de " & maxSteps & " pasos alcanzado.", vbExclamation
    Else
        UpdateVMStatus "COMPLETADO", "---", "Ejecuci�n finalizada"
        MsgBox "Ejecuci�n completada. Pasos ejecutados: " & stepCount, vbInformation
    End If
End Sub

Function ParseAndExecuteVM(instruction As String) As Boolean
    On Error GoTo ErrorHandler
    
    If InStr(instruction, ";") > 0 Then
        instruction = Trim(left(instruction, InStr(instruction, ";") - 1))
    End If
    
    If Trim(instruction) = "" Then
        CurrentInstructionVM = CurrentInstructionVM + 1
        ParseAndExecuteVM = True
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(Trim(instruction), " ")
    
    If UBound(parts) < 0 Then
        CurrentInstructionVM = CurrentInstructionVM + 1
        ParseAndExecuteVM = True
        Exit Function
    End If
    
    Dim Opcode As String: Opcode = UCase(Trim(parts(0)))
    Dim op1 As String: op1 = ""
    Dim op2 As String: op2 = ""
    
    If UBound(parts) >= 1 Then op1 = Trim(parts(1))
    If UBound(parts) >= 2 Then op2 = Trim(parts(2))
    
    Select Case Opcode
        Case "MOV": ParseAndExecuteVM = ExecuteMOV_VM(op1, op2)
        Case "ADD": ParseAndExecuteVM = ExecuteADD_VM(op1, op2)
        Case "SUB": ParseAndExecuteVM = ExecuteSUB_VM(op1, op2)
        Case "MUL": ParseAndExecuteVM = ExecuteMUL_VM(op1)
        Case "DIV": ParseAndExecuteVM = ExecuteDIV_VM(op1)
        Case "INC": ParseAndExecuteVM = ExecuteINC_VM(op1)
        Case "DEC": ParseAndExecuteVM = ExecuteDEC_VM(op1)
        Case "PUSH": ParseAndExecuteVM = ExecutePUSH_VM(op1)
        Case "POP": ParseAndExecuteVM = ExecutePOP_VM(op1)
        Case "CMP": ParseAndExecuteVM = ExecuteCMP_VM(op1, op2)
        Case "JMP": ParseAndExecuteVM = ExecuteJMP_VM(op1)
        Case "JZ", "JE": ParseAndExecuteVM = ExecuteJZ_VM(op1)
        Case "JNZ", "JNE": ParseAndExecuteVM = ExecuteJNZ_VM(op1)
        Case "HLT": IsRunning = False: ParseAndExecuteVM = True
        Case "NOP": CurrentInstructionVM = CurrentInstructionVM + 1: ParseAndExecuteVM = True
        Case Else: ParseAndExecuteVM = False
    End Select
    Exit Function
ErrorHandler:
    ParseAndExecuteVM = False
End Function

' ===================================================================================
' INSTRUCCIONES
' ===================================================================================

Function ExecuteMOV_VM(dest As String, src As String) As Boolean
    If SetRegisterValueVM(dest, GetOperandValueVM(src)) Then
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecuteMOV_VM = True
    Else
        ExecuteMOV_VM = False
    End If
End Function

Function ExecuteADD_VM(dest As String, src As String) As Boolean
    Dim Result As Long
    Result = GetRegisterValueVM(dest) + GetOperandValueVM(src)
    If SetRegisterValueVM(dest, Result) Then
        UpdateFlagsVM Result
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecuteADD_VM = True
    Else
        ExecuteADD_VM = False
    End If
End Function

Function ExecuteSUB_VM(dest As String, src As String) As Boolean
    Dim Result As Long
    Result = GetRegisterValueVM(dest) - GetOperandValueVM(src)
    If SetRegisterValueVM(dest, Result) Then
        UpdateFlagsVM Result
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecuteSUB_VM = True
    Else
        ExecuteSUB_VM = False
    End If
End Function

Function ExecuteMUL_VM(operand As String) As Boolean
    AX = AX * GetOperandValueVM(operand)
    UpdateFlagsVM AX
    CurrentInstructionVM = CurrentInstructionVM + 1
    ExecuteMUL_VM = True
End Function

Function ExecuteDIV_VM(operand As String) As Boolean
    Dim divisor As Long: divisor = GetOperandValueVM(operand)
    If divisor = 0 Then
        MsgBox "Error: Divisi�n por cero", vbCritical
        ExecuteDIV_VM = False
    Else
        AX = AX \ divisor
        UpdateFlagsVM AX
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecuteDIV_VM = True
    End If
End Function

Function ExecuteINC_VM(operand As String) As Boolean
    Dim Result As Long: Result = GetRegisterValueVM(operand) + 1
    If SetRegisterValueVM(operand, Result) Then
        UpdateFlagsVM Result
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecuteINC_VM = True
    Else
        ExecuteINC_VM = False
    End If
End Function

Function ExecuteDEC_VM(operand As String) As Boolean
    Dim Result As Long: Result = GetRegisterValueVM(operand) - 1
    If SetRegisterValueVM(operand, Result) Then
        UpdateFlagsVM Result
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecuteDEC_VM = True
    Else
        ExecuteDEC_VM = False
    End If
End Function

Function ExecutePUSH_VM(operand As String) As Boolean
    If StackPointerVM > 0 Then
        StackPointerVM = StackPointerVM - 1
        WriteVMMemory StackPointerVM, CStr(GetOperandValueVM(operand)), "STACK"
        SP = StackPointerVM
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecutePUSH_VM = True
    Else
        ExecutePUSH_VM = False
    End If
End Function

Function ExecutePOP_VM(dest As String) As Boolean
    If StackPointerVM < VM_SIZE - 1 Then
        Dim value As String: value = ReadVMMemory(StackPointerVM)
        StackPointerVM = StackPointerVM + 1
        SP = StackPointerVM
        If SetRegisterValueVM(dest, val(value)) Then
            CurrentInstructionVM = CurrentInstructionVM + 1
            ExecutePOP_VM = True
        Else
            ExecutePOP_VM = False
        End If
    Else
        ExecutePOP_VM = False
    End If
End Function

Function ExecuteCMP_VM(op1 As String, op2 As String) As Boolean
    UpdateFlagsVM (GetOperandValueVM(op1) - GetOperandValueVM(op2))
    CurrentInstructionVM = CurrentInstructionVM + 1
    ExecuteCMP_VM = True
End Function

Function ExecuteJMP_VM(address As String) As Boolean
    Dim addr As Integer: addr = val(address)
    If addr >= 0 And addr < VM_SIZE Then
        CurrentInstructionVM = addr
        ExecuteJMP_VM = True
    Else
        ExecuteJMP_VM = False
    End If
End Function

Function ExecuteJZ_VM(address As String) As Boolean
    If (flags And 1) = 1 Then
        ExecuteJZ_VM = ExecuteJMP_VM(address)
    Else
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecuteJZ_VM = True
    End If
End Function

Function ExecuteJNZ_VM(address As String) As Boolean
    If (flags And 1) = 0 Then
        ExecuteJNZ_VM = ExecuteJMP_VM(address)
    Else
        CurrentInstructionVM = CurrentInstructionVM + 1
        ExecuteJNZ_VM = True
    End If
End Function

' ===================================================================================
' AUXILIARES
' ===================================================================================

Function GetRegisterValueVM(regName As String) As Long
    Select Case UCase(regName)
        Case "AX": GetRegisterValueVM = AX
        Case "BX": GetRegisterValueVM = BX
        Case "CX": GetRegisterValueVM = CX
        Case "DX": GetRegisterValueVM = DX
        Case "SI": GetRegisterValueVM = SI
        Case "DI": GetRegisterValueVM = DI
        Case "BP": GetRegisterValueVM = BP
        Case "SP": GetRegisterValueVM = SP
        Case Else: GetRegisterValueVM = 0
    End Select
End Function

Function SetRegisterValueVM(regName As String, value As Long) As Boolean
    Select Case UCase(regName)
        Case "AX": AX = value
        Case "BX": BX = value
        Case "CX": CX = value
        Case "DX": DX = value
        Case "SI": SI = value
        Case "DI": DI = value
        Case "BP": BP = value
        Case "SP": SP = value: StackPointerVM = value
        Case Else: SetRegisterValueVM = False: Exit Function
    End Select
    SetRegisterValueVM = True
End Function

Function GetOperandValueVM(operand As String) As Long
    If left(operand, 1) = "[" And Right(operand, 1) = "]" Then
        GetOperandValueVM = val(ReadVMMemory(val(Mid(operand, 2, Len(operand) - 2))))
        Exit Function
    End If
    Dim regVal As Long: regVal = GetRegisterValueVM(operand)
    If regVal = 0 And IsNumeric(operand) Then
        GetOperandValueVM = val(operand)
    Else
        GetOperandValueVM = regVal
    End If
End Function

Sub UpdateFlagsVM(value As Long)
    If value = 0 Then flags = flags Or 1 Else flags = flags And Not 1
    If value < 0 Then flags = flags Or 2 Else flags = flags And Not 2
End Sub

' ===================================================================================
' INTERFAZ
' ===================================================================================

Sub CreateVMInterface()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("MemoriaVirtual")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "MemoriaVirtual"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws
        .Range("B1").value = "SIMULADOR DE MEMORIA VIRTUAL"
        .Range("B1:J1").Merge
        .Range("B1").Font.Bold = True
        .Range("B1").Font.Size = 14
        
        .Cells(3, 1).value = "Dir"
        Dim i As Integer, j As Integer
        For i = 0 To VM_COLS - 1
            .Cells(3, i + 2).value = Hex(i)
        Next i
        
        .Range("A3").Resize(1, VM_COLS + 1).Font.Bold = True
        .Range("A3").Resize(1, VM_COLS + 1).Interior.Color = RGB(200, 200, 200)
        .Range("A3").Resize(1, VM_COLS + 1).HorizontalAlignment = xlCenter
        
        For i = 0 To VM_ROWS - 1
            .Cells(4 + i, 1).value = "0x" & Format(Hex(i * VM_COLS), "00")
            .Cells(4 + i, 1).Font.Bold = True
            .Cells(4 + i, 1).Interior.Color = RGB(200, 200, 200)
            
            For j = 0 To VM_COLS - 1
                With .Cells(4 + i, j + 2)
                    .value = "00"
                    .HorizontalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous
                    .Interior.Color = RGB(240, 240, 240)
                    .Font.Name = "Courier New"
                    .Font.Size = 9
                End With
            Next j
        Next i
        
        .Columns("A").ColumnWidth = 6
        .Columns("B:Q").ColumnWidth = 4.5
        
        .Range("S3").value = "REGISTROS"
        .Range("S3:U3").Merge
        .Range("S3").Font.Bold = True
        .Range("S3").HorizontalAlignment = xlCenter
        .Range("S3").Interior.Color = RGB(200, 200, 200)
        
        .Range("S4").value = "Reg"
        .Range("T4").value = "Dec"
        .Range("U4").value = "Hex"
        .Range("S4:U4").Font.Bold = True
        
        .Range("S13").value = "ESTADO"
        .Range("S13:U13").Merge
        .Range("S13").Font.Bold = True
        .Range("S13").HorizontalAlignment = xlCenter
        .Range("S13").Interior.Color = RGB(200, 200, 200)
        
        .Range("S14").value = "Estado:"
        .Range("S15").value = "Dir:"
        .Range("S16").value = "Instr:"
        
        ' Mostrar registros iniciales
        UpdateRegisterDisplayVM
    End With
    
    CreateVMControlButtons
End Sub

Sub UpdateVMDisplay()
    Dim ws As Worksheet
    Set ws = Worksheets("MemoriaVirtual")
    
    Dim displayData() As Variant
    ReDim displayData(1 To VM_ROWS, 1 To VM_COLS)
    
    Dim i As Integer, j As Integer, addr As Integer
    For i = 0 To VM_ROWS - 1
        For j = 0 To VM_COLS - 1
            addr = i * VM_COLS + j
            displayData(i + 1, j + 1) = VirtualMemory(addr).value
        Next j
    Next i
    
    ws.Range("B4").Resize(VM_ROWS, VM_COLS).value = displayData
    
    For i = 0 To VM_ROWS - 1
        For j = 0 To VM_COLS - 1
            addr = i * VM_COLS + j
            With ws.Cells(4 + i, j + 2)
                Select Case VirtualMemory(addr).dataType
                    Case "INSTR": .Interior.Color = RGB(200, 255, 200)
                    Case "DATA": .Interior.Color = RGB(200, 220, 255)
                    Case "STACK": .Interior.Color = RGB(255, 240, 200)
                    Case Else: .Interior.Color = RGB(240, 240, 240)
                End Select
                If VirtualMemory(addr).Accessed Then .Font.Bold = True
            End With
        Next j
    Next i
End Sub

Sub HighlightCurrentInstructionVM()
    Dim ws As Worksheet
    Set ws = Worksheets("MemoriaVirtual")
    Dim addr As Integer: addr = CurrentInstructionVM
    Dim row As Integer: row = 4 + (addr \ VM_COLS)
    Dim col As Integer: col = 2 + (addr Mod VM_COLS)
    ws.Cells(row, col).Interior.Color = RGB(255, 255, 0)
End Sub

Sub UpdateRegisterDisplayVM()
    Dim ws As Worksheet
    Set ws = Worksheets("MemoriaVirtual")
    ws.Cells(5, 19).value = "AX": ws.Cells(5, 20).value = AX: ws.Cells(5, 21).value = "0x" & Hex(AX)
    ws.Cells(6, 19).value = "BX": ws.Cells(6, 20).value = BX: ws.Cells(6, 21).value = "0x" & Hex(BX)
    ws.Cells(7, 19).value = "CX": ws.Cells(7, 20).value = CX: ws.Cells(7, 21).value = "0x" & Hex(CX)
    ws.Cells(8, 19).value = "DX": ws.Cells(8, 20).value = DX: ws.Cells(8, 21).value = "0x" & Hex(DX)
    ws.Cells(9, 19).value = "SI": ws.Cells(9, 20).value = SI: ws.Cells(9, 21).value = "0x" & Hex(SI)
    ws.Cells(10, 19).value = "DI": ws.Cells(10, 20).value = DI: ws.Cells(10, 21).value = "0x" & Hex(DI)
    ws.Cells(11, 19).value = "SP": ws.Cells(11, 20).value = SP: ws.Cells(11, 21).value = "0x" & Hex(SP)
    ws.Cells(12, 19).value = "FL": ws.Cells(12, 20).value = flags: ws.Cells(12, 21).value = "0x" & Hex(flags)
End Sub

Sub UpdateVMStatus(status As String, address As String, details As String)
    Dim ws As Worksheet
    Set ws = Worksheets("MemoriaVirtual")
    ws.Range("T14").value = status
    ws.Range("T15").value = address
    ws.Range("T16").value = left(details, 30)
    Select Case status
        Case "EJECUTANDO": ws.Range("T14").Interior.Color = RGB(255, 255, 0)
        Case "COMPLETADO": ws.Range("T14").Interior.Color = RGB(0, 255, 0)
        Case "ERROR": ws.Range("T14").Interior.Color = RGB(255, 0, 0)
        Case Else: ws.Range("T14").Interior.Color = RGB(255, 255, 255)
    End Select
End Sub

Sub CreateVMControlButtons()
    Dim btn As Button, ws As Worksheet
    Set ws = Worksheets("MemoriaVirtual")
    On Error Resume Next: ws.Buttons.Delete: On Error GoTo 0
    
    Set btn = ws.Buttons.Add(50, 350, 100, 25)
    btn.OnAction = "ExecuteNextInstructionVM"
    btn.Characters.text = "Siguiente"
    
    Set btn = ws.Buttons.Add(160, 350, 100, 25)
    btn.OnAction = "ExecuteFullProgramVM"
    btn.Characters.text = "Ejecutar Todo"
    
    Set btn = ws.Buttons.Add(270, 350, 80, 25)
    btn.OnAction = "ResetVirtualMemory"
    btn.Characters.text = "Reiniciar"
End Sub

Public Sub IniciarMemoriaVirtual()
    InitializeVirtualMemorySimulator
End Sub

Public Sub EjecutarMemoriaVirtualCompleta()
    ExecuteFullProgramVM
End Sub
