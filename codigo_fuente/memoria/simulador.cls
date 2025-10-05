VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "simulador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub ConfigurarInterfaz()
    With Me
        .Cells.Clear
        
        ' Título
        .Range("A1:Z1").Merge
        .Range("A1").Value = "SIMULADOR x86 – ARQUITECTURA EN TIEMPO REAL"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(40, 80, 120)
        .Range("A1").Font.Color = RGB(255, 255, 255)
        
        ' === REGISTROS (izquierda) ===
        .Range("A3").Value = "REGISTROS"
        .Range("A3").Font.Bold = True
        .Range("A3").Font.Color = RGB(0, 80, 0)
        Dim regs As Variant: regs = Array("EAX", "EBX", "ECX", "EDX", "ESI", "EDI", "EBP", "ESP")
        Dim i As Integer
        For i = 0 To 7
            .Cells(5 + i, "A").Value = regs(i)
            .Cells(5 + i, "B").Value = 0
        Next i
        .Range("A5:B12").Borders.LineStyle = xlContinuous
        .Range("A5:B12").Interior.Color = RGB(240, 255, 240)
        
        ' === FLAGS (centro-izquierda) ===
        .Range("D3").Value = "FLAGS"
        .Range("D3").Font.Bold = True
        .Range("D3").Font.Color = RGB(120, 0, 0)
        Dim flags As Variant: flags = Array("ZF", "SF", "CF", "OF", "PF", "AF")
        For i = 0 To 5
            .Cells(5 + i, "D").Value = flags(i)
            .Cells(5 + i, "E").Value = "0"
        Next i
        .Range("D5:E10").Borders.LineStyle = xlContinuous
        .Range("D5:E10").Interior.Color = RGB(255, 240, 240)
        
        ' === ALU (centro-derecha) ===
        .Range("G3").Value = "ALU"
        .Range("G3").Font.Bold = True
        .Range("G3").Font.Color = RGB(100, 50, 0)
        .Range("G5").Value = "Operación:"
        .Range("G6").Value = "Op1:"
        .Range("G7").Value = "Op2:"
        .Range("G8").Value = "Resultado:"
        .Range("H5:H8").Value = "-"
        .Range("G5:H8").Borders.LineStyle = xlContinuous
        .Range("G5:H8").Interior.Color = RGB(255, 250, 240)
        
        ' === UNIDAD DE CONTROL (derecha) ===
        .Range("K3").Value = "UNIDAD DE CONTROL"
        .Range("K3").Font.Bold = True
        .Range("K3").Font.Color = RGB(80, 0, 80)
        .Range("K5").Value = "Etapa:"
        .Range("L5").Value = "Idle"
        .Range("K6").Value = "Instrucción:"
        .Range("L6").Value = "-"
        .Range("K5:L6").Borders.LineStyle = xlContinuous
        .Range("K5:L6").Interior.Color = RGB(250, 240, 255)
        
        ' BOTONES CON NOMBRES COMPLETOS DE MACROS
        Dim btnPaso As Button, btnTodo As Button, btnReset As Button
        On Error Resume Next
        .Buttons.Delete
        On Error GoTo 0
        
        Set btnPaso = .Buttons.Add(50, 250, 100, 30)
        With btnPaso
            .Name = "btnPaso"
            .Caption = "? Paso"
            ' USAR NOMBRE COMPLETO CON NOMBRE DE HOJA
            .OnAction = "'simulador.Simulador_Paso'"
            .Font.Bold = True
        End With
        
        Set btnTodo = .Buttons.Add(170, 250, 100, 30)
        With btnTodo
            .Name = "btnTodo"
            .Caption = "? Todo"
            ' USAR NOMBRE COMPLETO CON NOMBRE DE HOJA
            .OnAction = "'simulador.Simulador_Todo'"
            .Font.Bold = True
        End With
        
        Set btnReset = .Buttons.Add(290, 250, 100, 30)
        With btnReset
            .Name = "btnReset"
            .Caption = "?? Reiniciar"
            ' USAR NOMBRE COMPLETO CON NOMBRE DE HOJA
            .OnAction = "'simulador.Simulador_Reiniciar'"
            .Font.Bold = True
        End With
        
        .Columns("A:Z").ColumnWidth = 12
        .Rows("1:30").RowHeight = 20
    End With
End Sub
Sub Simulador_Paso()
    On Error GoTo ErrorHandler
    ModuloSimulador.EjecutarPasoAPaso
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en ejecución paso a paso: " & Err.Description, vbCritical
End Sub
Sub Simulador_Todo()
    On Error GoTo ErrorHandler
    ModuloSimulador.EjecutarProgramaCompleto
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en ejecución completa: " & Err.Description, vbCritical
End Sub
Sub Simulador_Reiniciar()
    On Error GoTo ErrorHandler
    ModuloSimulador.ReiniciarTodo
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al reiniciar: " & Err.Description, vbCritical
End Sub
