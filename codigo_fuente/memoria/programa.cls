VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "programa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub ConfigurarHojaPrograma()
    With Me
        .Cells.Clear
        
        ' Título
        .Range("A1:J1").Merge
        .Range("A1").Value = "INGRESA TU PROGRAMA EN ENSAMBLADOR x86"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(60, 100, 160)
        .Range("A1").Font.Color = RGB(255, 255, 255)
        
        ' Instrucciones de ejemplo
        .Range("A3").Value = "Ejemplo (usa Alt+Enter para saltos de línea):"
        .Range("A4").Value = "mov eax, 10" & Chr(10) & "mov ebx, 3" & Chr(10) & "add eax, ebx"
        .Range("A4").Font.Italic = True
        .Range("A4").Font.Size = 9
        .Range("A4").Font.Color = RGB(80, 80, 80)
        .Range("A4").WrapText = True
        
        ' Cuadro de código
        .Range("A6:J20").Merge
        .Range("A6").Borders.LineStyle = xlContinuous
        .Range("A6").Interior.Color = RGB(250, 250, 255)
        .Range("A6").Font.Name = "Courier New"
        .Range("A6").Font.Size = 10
        .Range("A6").WrapText = True
        .Range("A6").VerticalAlignment = xlTop
        
        ' BOTONES CON NOMBRES COMPLETOS DE MACROS
        Dim btnEjecutar As Button, btnReiniciar As Button
        On Error Resume Next
        .Buttons.Delete
        On Error GoTo 0
        
        Set btnEjecutar = .Buttons.Add(100, 350, 120, 30)
        With btnEjecutar
            .Name = "btnEjecutar"
            .Caption = "? Ejecutar"
            .OnAction = "'programa.Programa_Ejecutar'"
            .Font.Bold = True
        End With
        
        Set btnReiniciar = .Buttons.Add(250, 350, 120, 30)
        With btnReiniciar
            .Name = "btnReiniciar"
            .Caption = "?? Reiniciar"
            .OnAction = "'programa.Programa_Reiniciar'"
            .Font.Bold = True
        End With
        
        .Columns("A:J").ColumnWidth = 12
        .Rows(6).RowHeight = 60
    End With
End Sub
Sub Programa_Ejecutar()
    On Error GoTo ErrorHandler
    ' Cargar programa y ejecutar
    ModuloSimulador.CargarProgramaDesdeHoja
    Sheets("Simulador").Activate
    ModuloSimulador.EjecutarProgramaCompleto
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al ejecutar: " & Err.Description, vbCritical
End Sub
Sub Programa_Reiniciar()
    On Error GoTo ErrorHandler
    ModuloSimulador.ReiniciarTodo
    Sheets("Simulador").Activate
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al reiniciar: " & Err.Description, vbCritical
End Sub
