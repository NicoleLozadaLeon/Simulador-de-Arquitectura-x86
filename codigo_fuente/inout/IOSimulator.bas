Attribute VB_Name = "Módulo1"
' Módulo: IOSimulator
' Descripción: Simulador completo de Entrada/Salida en UN SOLO MÓDULO
' Versión simplificada sin dependencias entre módulos

Option Explicit

' ============================================
' VARIABLES GLOBALES DEL SISTEMA
' ============================================
Public InputBuffer As String
Public OutputBuffer As String
Public MemoryData(0 To 255) As Byte
Public DataBusValue As String
Public AddressBusValue As String
Public ControlBusSignal As String
Public CurrentIOAddress As Integer
Public CurrentOperationType As String

' ============================================
' PROCEDIMIENTO PRINCIPAL - EJECUTAR ESTE
' ============================================
Sub IniciarSimuladorIO()
    On Error GoTo ErrorHandler
    
    ' Limpiar memoria
    LimpiarMemoria
    
    ' Crear hojas
    CrearHojaEntrada
    CrearHojaSalida
    CrearHojaMemoria
    
    ' Inicializar valores
    InputBuffer = ""
    OutputBuffer = ""
    DataBusValue = "00000000"
    AddressBusValue = "00000000"
    ControlBusSignal = "IDLE"
    CurrentIOAddress = 0
    CurrentOperationType = "Ninguna"
    
    ' Actualizar estado
    ActualizarMemoriaVisual
    
    MsgBox "Simulador de Entrada/Salida Iniciado" & vbCrLf & vbCrLf & _
           "HOJAS CREADAS:" & vbCrLf & _
           "- INPUT: Dispositivos de Entrada (Teclado)" & vbCrLf & _
           "- OUTPUT: Dispositivos de Salida (Pantalla)" & vbCrLf & _
           "- MEMORIA_IO: Memoria y Buses del Sistema" & vbCrLf & vbCrLf & _
           "Use el teclado virtual en la hoja INPUT.", _
           vbInformation, "Sistema E/S Listo"
    
    Worksheets("INPUT").Activate
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al iniciar: " & Err.Description, vbCritical
End Sub

' ============================================
' GESTIÓN DE MEMORIA
' ============================================
Sub LimpiarMemoria()
    Dim i As Integer
    For i = 0 To 255
        MemoryData(i) = 0
    Next i
End Sub

' ============================================
' CREACIÓN DE HOJA INPUT (TECLADO)
' ============================================
Sub CrearHojaEntrada()
    Dim ws As Worksheet
    Dim i As Integer, row As Integer, col As Integer
    
    ' Eliminar si existe
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("INPUT").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Crear nueva hoja
    Set ws = Worksheets.Add(Before:=Worksheets(1))
    ws.Name = "INPUT"
    
    With ws
        ' Título
        .Range("B2:P3").Merge
        .Range("B2").value = "UNIDAD DE ENTRADA - TECLADO"
        With .Range("B2")
            .Font.Bold = True
            .Font.Size = 18
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(70, 130, 180)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' Información del dispositivo
        .Range("B5").value = "DISPOSITIVO:"
        .Range("B5").Font.Bold = True
        .Range("C5").value = "Teclado (Keyboard) - Dirección E/S: 0x60"
        
        ' Título del teclado
        .Range("B10:P10").Merge
        .Range("B10").value = "TECLADO VIRTUAL - Haga clic en las teclas"
        With .Range("B10")
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(100, 100, 100)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' Crear teclado
        row = 12
        
        ' Fila números
        Dim numeros As Variant
        numeros = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "0")
        col = 2
        For i = 0 To 9
            CrearTecla ws, row, col + (i * 2), CStr(numeros(i))
        Next i
        
        ' Fila QWERTY
        row = row + 3
        Dim qwerty As Variant
        qwerty = Array("Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P")
        col = 3
        For i = 0 To 9
            CrearTecla ws, row, col + (i * 2), CStr(qwerty(i))
        Next i
        
        ' Fila ASDF
        row = row + 3
        Dim asdf As Variant
        asdf = Array("A", "S", "D", "F", "G", "H", "J", "K", "L")
        col = 4
        For i = 0 To 8
            CrearTecla ws, row, col + (i * 2), CStr(asdf(i))
        Next i
        
        ' Fila ZXCV
        row = row + 3
        Dim zxcv As Variant
        zxcv = Array("Z", "X", "C", "V", "B", "N", "M")
        col = 5
        For i = 0 To 6
            CrearTecla ws, row, col + (i * 2), CStr(zxcv(i))
        Next i
        
        ' Espacio
        row = row + 3
        .Range(.Cells(row, 6), .Cells(row + 1, 15)).Merge
        With .Range(.Cells(row, 6), .Cells(row + 1, 15))
            .value = "ESPACIO"
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
            .Name = "Key_SPACE"
        End With
        
        ' Buffer de entrada
        row = row + 4
        .Range(.Cells(row, 2), .Cells(row, 16)).Merge
        .Cells(row, 2).value = "BUFFER DE ENTRADA"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).HorizontalAlignment = xlCenter
        .Cells(row, 2).Interior.Color = RGB(200, 200, 200)
        
        row = row + 1
        .Range(.Cells(row, 2), .Cells(row + 2, 16)).Merge
        .Cells(row, 2).Name = "InputBuffer"
        With .Cells(row, 2)
            .Font.Name = "Courier New"
            .Font.Size = 14
            .Interior.Color = RGB(255, 255, 220)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Estado
        row = row + 4
        .Cells(row, 2).value = "Última tecla:"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 3).value = "Ninguna"
        .Cells(row, 3).Name = "LastKey"
        
        row = row + 1
        .Cells(row, 2).value = "Código ASCII:"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 3).value = "00h"
        .Cells(row, 3).Name = "ASCIICode"
        
        ' Botones
        row = row + 2
        CrearBoton ws, row, 2, "Enviar a Memoria", "EnviarAMemoria"
        CrearBoton ws, row, 5, "Limpiar", "LimpiarBufferEntrada"
        CrearBoton ws, row, 8, "Ver Memoria", "IrAMemoria"
    End With
    
    ' Asignar macros a teclas
    AsignarMacrosTeclas
End Sub

Sub CrearTecla(ws As Worksheet, row As Integer, col As Integer, tecla As String)
    With ws.Range(ws.Cells(row, col), ws.Cells(row + 1, col + 1))
        .Merge
        .value = tecla
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
        .Name = "Key_" & tecla
    End With
End Sub

Sub AsignarMacrosTeclas()
    Dim ws As Worksheet
    Set ws = Worksheets("INPUT")
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        If left(shp.Name, 4) = "Key_" Then
            shp.OnAction = "PresionarTecla"
        End If
    Next shp
End Sub

Sub PresionarTecla()
    Dim ws As Worksheet
    Set ws = Worksheets("INPUT")
    
    Dim nombreShape As String
    nombreShape = Application.Caller
    
    Dim tecla As String
    tecla = Mid(nombreShape, 5) ' Quitar "Key_"
    
    If tecla = "SPACE" Then tecla = " "
    
    ' Agregar al buffer
    InputBuffer = InputBuffer & tecla
    ws.Range("InputBuffer").value = InputBuffer
    
    ' Actualizar estado
    Dim asciiCode As Integer
    asciiCode = Asc(tecla)
    ws.Range("LastKey").value = tecla
    ws.Range("ASCIICode").value = Format(Hex(asciiCode), "00") & "h (" & asciiCode & ")"
    
    ' Actualizar buses
    DataBusValue = DecimalABinario(asciiCode, 8)
    AddressBusValue = "01100000" ' 0x60
    ControlBusSignal = "READ"
    CurrentIOAddress = &H60
    CurrentOperationType = "Lectura de Teclado"
    
    ActualizarMemoriaVisual
End Sub

' ============================================
' CREACIÓN DE HOJA OUTPUT (PANTALLA)
' ============================================
Sub CrearHojaSalida()
    Dim ws As Worksheet
    Dim row As Integer
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("OUTPUT").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.count))
    ws.Name = "OUTPUT"
    
    With ws
        ' Título
        .Range("B2:P3").Merge
        .Range("B2").value = "UNIDAD DE SALIDA - PANTALLA"
        With .Range("B2")
            .Font.Bold = True
            .Font.Size = 18
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(220, 20, 60)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' Información
        .Range("B5").value = "DISPOSITIVO:"
        .Range("B5").Font.Bold = True
        .Range("C5").value = "Monitor/Pantalla - Dirección E/S: 0x3D4"
        
        ' Pantalla
        .Range("B10:P10").Merge
        .Range("B10").value = "PANTALLA VIRTUAL"
        With .Range("B10")
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(50, 50, 50)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        row = 12
        .Range(.Cells(row, 2), .Cells(row + 20, 17)).Merge
        .Cells(row, 2).Name = "ScreenDisplay"
        With .Cells(row, 2)
            .Font.Name = "Consolas"
            .Font.Size = 12
            .Font.Color = RGB(0, 255, 0)
            .Interior.Color = RGB(0, 0, 0)
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlTop
            .WrapText = True
        End With
        
        ' Buffer de salida
        row = row + 22
        .Range(.Cells(row, 2), .Cells(row, 17)).Merge
        .Cells(row, 2).value = "BUFFER DE SALIDA (VRAM)"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).HorizontalAlignment = xlCenter
        .Cells(row, 2).Interior.Color = RGB(200, 200, 200)
        
        row = row + 1
        .Range(.Cells(row, 2), .Cells(row + 2, 17)).Merge
        .Cells(row, 2).Name = "OutputBuffer"
        With .Cells(row, 2)
            .Font.Name = "Courier New"
            .Interior.Color = RGB(255, 240, 220)
            .Borders.LineStyle = xlContinuous
            .WrapText = True
        End With
        
        ' Botones
        row = row + 4
        CrearBoton ws, row, 2, "Leer de Memoria", "LeerDeMemoria"
        CrearBoton ws, row, 5, "Mostrar en Pantalla", "MostrarEnPantalla"
        CrearBoton ws, row, 8, "Limpiar Pantalla", "LimpiarPantalla"
        CrearBoton ws, row, 11, "Ver Memoria", "IrAMemoria"
    End With
End Sub

' ============================================
' CREACIÓN DE HOJA MEMORIA
' ============================================
Sub CrearHojaMemoria()
    Dim ws As Worksheet
    Dim i As Integer, j As Integer, row As Integer
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("MEMORIA_IO").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.count))
    ws.Name = "MEMORIA_IO"
    
    With ws
        ' Título
        .Range("B2:S2").Merge
        .Range("B2").value = "MEMORIA Y BUSES DEL SISTEMA"
        With .Range("B2")
            .Font.Bold = True
            .Font.Size = 16
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(75, 0, 130)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' Memoria RAM
        row = 4
        .Range(.Cells(row, 2), .Cells(row, 9)).Merge
        .Cells(row, 2).value = "MEMORIA RAM (256 BYTES)"
        With .Cells(row, 2)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(70, 130, 180)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' Encabezados
        row = row + 1
        .Cells(row, 2).value = "Dir"
        .Cells(row, 2).Font.Bold = True
        For i = 0 To 7
            .Cells(row, 3 + i).value = Format(Hex(i), "0")
            .Cells(row, 3 + i).Font.Bold = True
            .Cells(row, 3 + i).HorizontalAlignment = xlCenter
            .Cells(row, 3 + i).Interior.Color = RGB(200, 200, 200)
        Next i
        
        ' Celdas de memoria
        For i = 0 To 15
            row = row + 1
            .Cells(row, 2).value = "0x" & Format(Hex(i * 8), "00")
            .Cells(row, 2).Font.Name = "Courier New"
            .Cells(row, 2).Font.Bold = True
            .Cells(row, 2).Interior.Color = RGB(220, 220, 220)
            
            For j = 0 To 7
                With .Cells(row, 3 + j)
                    .value = "00"
                    .Font.Name = "Courier New"
                    .Font.Size = 9
                    .HorizontalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous
                    .Name = "Mem_" & Format(i * 8 + j, "000")
                End With
            Next j
        Next i
        
        ' BUSES
        Dim colBus As Integer
        colBus = 12
        row = 4
        
        ' Bus de Datos
        .Range(.Cells(row, colBus), .Cells(row, colBus + 8)).Merge
        .Cells(row, colBus).value = "BUS DE DATOS (8 BITS)"
        With .Cells(row, colBus)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(100, 149, 237)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        row = row + 1
        .Range(.Cells(row, colBus), .Cells(row + 2, colBus + 8)).Merge
        .Cells(row, colBus).Name = "DataBus"
        With .Cells(row, colBus)
            .value = "00000000"
            .Font.Name = "Courier New"
            .Font.Size = 16
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(173, 216, 230)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Bus de Direcciones
        row = row + 4
        .Range(.Cells(row, colBus), .Cells(row, colBus + 8)).Merge
        .Cells(row, colBus).value = "BUS DE DIRECCIONES (8 BITS)"
        With .Cells(row, colBus)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(50, 205, 50)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        row = row + 1
        .Range(.Cells(row, colBus), .Cells(row + 2, colBus + 8)).Merge
        .Cells(row, colBus).Name = "AddressBus"
        With .Cells(row, colBus)
            .value = "00000000"
            .Font.Name = "Courier New"
            .Font.Size = 16
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(144, 238, 144)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Bus de Control
        row = row + 4
        .Range(.Cells(row, colBus), .Cells(row, colBus + 8)).Merge
        .Cells(row, colBus).value = "BUS DE CONTROL"
        With .Cells(row, colBus)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(220, 20, 60)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        row = row + 1
        .Range(.Cells(row, colBus), .Cells(row + 2, colBus + 8)).Merge
        .Cells(row, colBus).Name = "ControlBus"
        With .Cells(row, colBus)
            .value = "IDLE"
            .Font.Name = "Courier New"
            .Font.Size = 14
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(255, 182, 193)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Estado del sistema
        row = row + 4
        .Range(.Cells(row, colBus), .Cells(row, colBus + 8)).Merge
        .Cells(row, colBus).value = "ESTADO DEL SISTEMA"
        With .Cells(row, colBus)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(75, 0, 130)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        row = row + 1
        .Cells(row, colBus).value = "Operación:"
        .Cells(row, colBus).Font.Bold = True
        .Range(.Cells(row, colBus + 1), .Cells(row, colBus + 4)).Merge
        .Cells(row, colBus + 1).Name = "CurrentOp"
        .Cells(row, colBus + 1).value = "Ninguna"
        
        row = row + 1
        .Cells(row, colBus).value = "Dirección:"
        .Cells(row, colBus).Font.Bold = True
        .Cells(row, colBus + 1).Name = "IOAddr"
        .Cells(row, colBus + 1).value = "0x00"
        
        ' Botones
        row = row + 3
        CrearBoton ws, row, 2, "Actualizar Vista", "ActualizarMemoriaVisual"
        CrearBoton ws, row, 5, "Limpiar Memoria", "LimpiarTodo"
        CrearBoton ws, row, 8, "Ver INPUT", "IrAInput"
        CrearBoton ws, row, 11, "Ver OUTPUT", "IrAOutput"
    End With
End Sub

Sub CrearBoton(ws As Worksheet, row As Integer, col As Integer, texto As String, macro As String)
    Dim btn As Button
    Dim leftPos As Double, topPos As Double
    
    leftPos = ws.Cells(row, col).left
    topPos = ws.Cells(row, col).top
    
    Set btn = ws.Buttons.Add(leftPos, topPos, 100, 25)
    btn.OnAction = macro
    btn.Characters.text = texto
    btn.Font.Bold = True
End Sub

' ============================================
' OPERACIONES DEL SISTEMA
' ============================================
Sub EnviarAMemoria()
    If InputBuffer = "" Then
        MsgBox "El buffer está vacío", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To Len(InputBuffer)
        If i - 1 <= 255 Then
            MemoryData(i - 1) = Asc(Mid(InputBuffer, i, 1))
        End If
    Next i
    
    ActualizarMemoriaVisual
    Worksheets("MEMORIA_IO").Activate
    
    MsgBox "Datos enviados a memoria: " & InputBuffer, vbInformation
End Sub

Sub LeerDeMemoria()
    Dim texto As String
    Dim i As Integer
    
    texto = ""
    For i = 0 To 255
        If MemoryData(i) = 0 Then Exit For
        texto = texto & Chr(MemoryData(i))
    Next i
    
    If texto = "" Then
        MsgBox "No hay datos en memoria", vbInformation
        Exit Sub
    End If
    
    OutputBuffer = texto
    Worksheets("OUTPUT").Range("OutputBuffer").value = texto
    
    ' Actualizar buses
    DataBusValue = DecimalABinario(Len(texto), 8)
    AddressBusValue = "11011000" ' 0x3D4
    ControlBusSignal = "WRITE"
    CurrentIOAddress = &H3D4
    CurrentOperationType = "Escritura a Pantalla"
    
    ActualizarMemoriaVisual
    
    MsgBox "Datos leídos de memoria: " & texto, vbInformation
End Sub

Sub MostrarEnPantalla()
    If OutputBuffer = "" Then
        MsgBox "El buffer de salida está vacío", vbInformation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Worksheets("OUTPUT")
    
    Dim i As Integer
    Dim textoActual As String
    
    textoActual = ""
    ws.Range("ScreenDisplay").value = ""
    
    For i = 1 To Len(OutputBuffer)
        textoActual = textoActual & Mid(OutputBuffer, i, 1)
        ws.Range("ScreenDisplay").value = textoActual
        Esperar 0.05
    Next i
    
    MsgBox "Texto mostrado en pantalla", vbInformation
End Sub

Sub ActualizarMemoriaVisual()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Worksheets("MEMORIA_IO")
    
    If ws Is Nothing Then Exit Sub
    
    Dim i As Integer
    For i = 0 To 127
        ws.Range("Mem_" & Format(i, "000")).value = Format(Hex(MemoryData(i)), "00")
        
        If MemoryData(i) <> 0 Then
            ws.Range("Mem_" & Format(i, "000")).Interior.Color = RGB(255, 255, 102)
            ws.Range("Mem_" & Format(i, "000")).Font.Bold = True
        Else
            ws.Range("Mem_" & Format(i, "000")).Interior.Color = RGB(255, 255, 255)
            ws.Range("Mem_" & Format(i, "000")).Font.Bold = False
        End If
    Next i
    
    ' Actualizar buses
    ws.Range("DataBus").value = DataBusValue
    ws.Range("AddressBus").value = AddressBusValue
    ws.Range("ControlBus").value = ControlBusSignal
    ws.Range("CurrentOp").value = CurrentOperationType
    ws.Range("IOAddr").value = "0x" & Format(Hex(CurrentIOAddress), "00")
    
    ' Colorear bus de control
    Select Case ControlBusSignal
        Case "READ"
            ws.Range("ControlBus").Interior.Color = RGB(144, 238, 144)
        Case "WRITE"
            ws.Range("ControlBus").Interior.Color = RGB(255, 182, 193)
        Case Else
            ws.Range("ControlBus").Interior.Color = RGB(240, 240, 240)
    End Select
    
    On Error GoTo 0
End Sub

' ============================================
' UTILIDADES Y NAVEGACIÓN
' ============================================
Sub LimpiarBufferEntrada()
    InputBuffer = ""
    Worksheets("INPUT").Range("InputBuffer").value = ""
    Worksheets("INPUT").Range("LastKey").value = "Ninguna"
    Worksheets("INPUT").Range("ASCIICode").value = "00h"
End Sub

Sub LimpiarPantalla()
    OutputBuffer = ""
    Worksheets("OUTPUT").Range("ScreenDisplay").value = ""
    Worksheets("OUTPUT").Range("OutputBuffer").value = ""
End Sub

Sub LimpiarTodo()
    If MsgBox("¿Limpiar toda la memoria?", vbYesNo + vbQuestion) = vbYes Then
        LimpiarMemoria
        ActualizarMemoriaVisual
        MsgBox "Memoria limpiada", vbInformation
    End If
End Sub

Sub IrAInput()
    Worksheets("INPUT").Activate
End Sub

Sub IrAOutput()
    Worksheets("OUTPUT").Activate
End Sub

Sub IrAMemoria()
    Worksheets("MEMORIA_IO").Activate
End Sub

Function DecimalABinario(numero As Integer, bits As Integer) As String
    Dim binario As String
    Dim i As Integer
    
    binario = ""
    For i = bits - 1 To 0 Step -1
        If numero And (2 ^ i) Then
            binario = binario & "1"
        Else
            binario = binario & "0"
        End If
    Next i
    
    DecimalABinario = binario
End Function

Sub Esperar(segundos As Double)
    Dim inicio As Double
    inicio = Timer
    Do While Timer < inicio + segundos
        DoEvents
    Loop
End Sub

