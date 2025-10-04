VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HojaEnsamblador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private regEAX As Long, regEBX As Long, regECX As Long, regEDX As Long
Private instrucciones As Collection
Private instruccionActual As Long
Private Sub Worksheet_Activate()
    Call ConfigurarHoja
End Sub
Public Sub ConfigurarHoja()
    With Me
        .Cells.Clear
        
        ' === TÍTULO ===
        .Range("A1:N1").Merge
        .Range("A1").Value = "SIMULADOR DE C A ENSAMBLADOR x86"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(70, 130, 180)
        .Range("A1").Font.Color = RGB(255, 255, 255)
        
        ' === ENTRADA DE CÓDIGO C ===
        .Range("A3").Value = "Ingresa código C simple:"
        .Range("A3").Font.Bold = True
        .Range("A3").Font.Color = RGB(0, 80, 150)
        .Range("A4").Value = "// Ej: a = 5; b = a + 3;"
        .Range("A4").Font.Italic = True
        .Range("A4").Font.Size = 9
        .Range("A4").Font.Color = RGB(100, 100, 100)
        .Range("A5:D8").Merge
        .Range("A5").Borders.LineStyle = xlContinuous
        .Range("A5").Interior.Color = RGB(240, 248, 255)
        .Range("A5").HorizontalAlignment = xlLeft
        .Range("A5").VerticalAlignment = xlTop
        .Range("A5").WrapText = True
        
        ' === BOTÓN: TRADUCIR ===
        Dim btn1 As Button
        On Error Resume Next
        .Buttons("btnTraducir").Delete
        On Error GoTo 0
        Set btn1 = .Buttons.Add(50, 180, 140, 30)
        With btn1
            .Name = "btnTraducir"
            .OnAction = "TraducirDesdeHoja"
            .Caption = "Traducir a Ensamblador"
            .Font.Bold = True
        End With
        
        ' === SALIDA: ENSAMBLADOR ===
        .Range("F3").Value = "Código Ensamblador x86:"
        .Range("F3").Font.Bold = True
        .Range("F3").Font.Color = RGB(0, 100, 0)
        .Range("F5:I15").Merge
        .Range("F5").Borders.LineStyle = xlContinuous
        .Range("F5").Interior.Color = RGB(250, 250, 230)
        .Range("F5").Font.Name = "Courier New"
        .Range("F5").Font.Size = 10
        .Range("F5").HorizontalAlignment = xlLeft
        .Range("F5").VerticalAlignment = xlTop
        .Range("F5").WrapText = True
        
        ' === BOTÓN: EJECUTAR ===
        On Error Resume Next
        .Buttons("btnEjecutar").Delete
        On Error GoTo 0
        Dim btn2 As Button
        Set btn2 = .Buttons.Add(450, 180, 140, 30)
        With btn2
            .Name = "btnEjecutar"
            .OnAction = "EjecutarPasoDesdeHoja"
            .Caption = "Ejecutar Paso a Paso"
            .Font.Bold = True
        End With
        
        ' === REGISTROS ===
        .Range("K3").Value = "Registros (Estado):"
        .Range("K3").Font.Bold = True
        .Range("K3").Font.Color = RGB(150, 0, 0)
        .Range("K5").Value = "EAX:": .Range("L5").Value = 0
        .Range("K6").Value = "EBX:": .Range("L6").Value = 0
        .Range("K7").Value = "ECX:": .Range("L7").Value = 0
        .Range("K8").Value = "EDX:": .Range("L8").Value = 0
        .Range("K5:L8").Borders.LineStyle = xlContinuous
        .Range("K5:L8").Interior.Color = RGB(255, 245, 245)
        
        ' === INSTRUCCIÓN ACTUAL ===
        .Range("K10").Value = "Instrucción:"
        .Range("L10:N12").Merge
        .Range("L10").Interior.Color = RGB(230, 230, 250)
        .Range("L10").Borders.LineStyle = xlContinuous
        .Range("L10").Value = "(Ninguna)"
        
        .Columns("A:N").ColumnWidth = 12
    End With
End Sub
Public Property Get RegistroEAX() As Long: RegistroEAX = regEAX: End Property
Public Property Get RegistroEBX() As Long: RegistroEBX = regEBX: End Property
Public Property Get RegistroECX() As Long: RegistroECX = regECX: End Property
Public Property Get RegistroEDX() As Long: RegistroEDX = regEDX: End Property
Public Sub ActualizarRegistrosEnHoja()
    Me.Range("L5").Value = regEAX
    Me.Range("L6").Value = regEBX
    Me.Range("L7").Value = regECX
    Me.Range("L8").Value = regEDX
End Sub
Public Sub MostrarInstruccion(lineaASM As String)
    Me.Range("L10").Value = lineaASM
End Sub
Public Function ObtenerCodigoC() As String
    ObtenerCodigoC = Trim(Me.Range("A5").Value)
End Function
Public Sub MostrarEnsamblador(codigo As String)
    Me.Range("F5").Value = codigo
End Sub
Public Sub ReiniciarSimulacion()
    regEAX = 0: regEBX = 0: regECX = 0: regEDX = 0
    Set instrucciones = New Collection
    instruccionActual = 0
    ActualizarRegistrosEnHoja
    Me.Range("L10").Value = "(Listo para ejecutar)"
End Sub
Public Sub AgregarInstrucciones(lineas() As String)
    Dim i As Long
    For i = 0 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            instrucciones.Add Trim(lineas(i))
        End If
    Next i
End Sub
Public Function ObtenerSiguienteInstruccion() As String
    If instrucciones Is Nothing Or instruccionActual >= instrucciones.Count Then
        ObtenerSiguienteInstruccion = ""
    Else
        instruccionActual = instruccionActual + 1
        ObtenerSiguienteInstruccion = instrucciones(instruccionActual)
    End If
End Function
Public Function HayMasInstrucciones() As Boolean
    If instrucciones Is Nothing Then
        HayMasInstrucciones = False
    Else
        HayMasInstrucciones = (instruccionActual < instrucciones.Count)
    End If
End Function
Public Sub EjecutarInstruccion(lineaASM As String)
    Dim partes() As String
    Dim destino As String, fuente As String
    Dim valor As Long
    Dim lineaMin As String
    
    lineaMin = LCase(Trim(lineaASM))
    
    If InStr(1, lineaMin, "mov") > 0 Then
        partes = Split(Replace(lineaMin, "mov ", ""), ",")
        destino = Trim(partes(0))
        fuente = Trim(partes(1))
        
        If IsNumeric(fuente) Then
            valor = CLng(fuente)
        ElseIf fuente = "eax" Then
            valor = regEAX
        ElseIf fuente = "ebx" Then
            valor = regEBX
        ElseIf InStr(fuente, "[a]") > 0 Then
            valor = regEAX
        ElseIf InStr(fuente, "[b]") > 0 Then
            valor = regEBX
        ElseIf InStr(fuente, "[c]") > 0 Then
            valor = regECX
        End If
        
        If destino = "eax" Then regEAX = valor
        If destino = "ebx" Then regEBX = valor
        If destino = "ecx" Then regECX = valor
        If destino = "edx" Then regEDX = valor
        If InStr(destino, "[a]") > 0 Then regEAX = valor
        If InStr(destino, "[b]") > 0 Then regEBX = valor
        If InStr(destino, "[c]") > 0 Then regECX = valor

    ElseIf InStr(1, lineaMin, "add") > 0 Then
        partes = Split(Replace(lineaMin, "add ", ""), ",")
        destino = Trim(partes(0))
        fuente = Trim(partes(1))
        valor = IIf(IsNumeric(fuente), CLng(fuente), regEAX)
        If destino = "eax" Then regEAX = regEAX + valor

    ElseIf InStr(1, lineaMin, "sub") > 0 Then
        partes = Split(Replace(lineaMin, "sub ", ""), ",")
        destino = Trim(partes(0))
        fuente = Trim(partes(1))
        valor = IIf(IsNumeric(fuente), CLng(fuente), regEAX)
        If destino = "eax" Then regEAX = regEAX - valor
    End If
End Sub
