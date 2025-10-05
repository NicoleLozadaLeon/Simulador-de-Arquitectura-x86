Attribute VB_Name = "ModuloCiclo"
Option Explicit

Private InstruccionActual As String
Private opcode As String
Private operandos As String
Private Memoria(0 To MEM_SIZE - 1) As String
Public Sub EjecutarCiclo()
    Fetch
    If Len(InstruccionActual) > 0 Then
        Decode
        Execute
    End If
    ActualizarEIP
End Sub
Private Sub Fetch()
    If eip < MEM_SIZE Then
        InstruccionActual = Memoria(eip)
        eip = eip + 1
    Else
        InstruccionActual = ""
    End If
End Sub
Private Sub Decode()
    Dim partes() As String
    If Len(InstruccionActual) > 0 Then
        partes = Split(InstruccionActual, " ")
        If UBound(partes) >= 0 Then
            opcode = partes(0)
            If UBound(partes) >= 1 Then
                operandos = partes(1)
            Else
                operandos = ""
            End If
        Else
            opcode = ""
            operandos = ""
        End If
    Else
        opcode = ""
        operandos = ""
    End If
End Sub
Private Sub ActualizarEIP()
    If eip >= MEM_SIZE Then
        eip = 0
    End If
End Sub
Public Sub CargarPrograma(programa() As String)
    Dim i As Long
    For i = 0 To UBound(programa)
        If i < MEM_SIZE Then
            Memoria(i) = programa(i)
        End If
    Next i
    eip = 0
End Sub
Public Sub EjecutarProgramaCompleto()
    Do While eip < MEM_SIZE And Len(Memoria(eip)) > 0
        EjecutarCiclo
    Loop
End Sub
   Private Sub Execute()
       If Len(InstruccionActual) > 0 Then
           ParsearInstruccion InstruccionActual
       End If
   End Sub

