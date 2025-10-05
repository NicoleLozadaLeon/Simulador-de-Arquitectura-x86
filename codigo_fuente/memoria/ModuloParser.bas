Attribute VB_Name = "ModuloParser"
' ===== ModuloParser =====
Option Explicit
Public Sub ParsearYEjecutar(inst As String)
    Dim partes() As String
    partes = Split(Trim(inst), " ")
    If UBound(partes) < 0 Then Exit Sub
    
    Select Case UCase(partes(0))
        Case "MOV"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarMOV ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "ADD"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarADD ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "SUB"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarSUB ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "MUL"
            If UBound(partes) >= 1 Then
                ModuloAlu.EjecutarMUL ExtraerOperando(partes(1))
            End If
        Case "DIV"
            If UBound(partes) >= 1 Then
                ModuloAlu.EjecutarDIV ExtraerOperando(partes(1))
            End If
        Case "IMUL"
            If UBound(partes) >= 1 Then
                ModuloAlu.EjecutarIMUL ExtraerOperando(partes(1))
            End If
        Case "IDIV"
            If UBound(partes) >= 1 Then
                ModuloAlu.EjecutarIDIV ExtraerOperando(partes(1))
            End If
        Case "AND"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarAND ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "OR"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarOR ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "XOR"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarXOR ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "NOT"
            If UBound(partes) >= 1 Then
                ModuloAlu.EjecutarNOT ExtraerOperando(partes(1))
            End If
        Case "CMP"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarCMP ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "TEST"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarTEST ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "INC"
            If UBound(partes) >= 1 Then
                ModuloAlu.EjecutarINC ExtraerOperando(partes(1))
            End If
        Case "DEC"
            If UBound(partes) >= 1 Then
                ModuloAlu.EjecutarDEC ExtraerOperando(partes(1))
            End If
        Case "SHL"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarSHL ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "SHR"
            If UBound(partes) >= 2 Then
                ModuloAlu.EjecutarSHR ExtraerOperando(partes(1)), ExtraerOperando(partes(2))
            End If
        Case "NOP"
        Case "HLT"
            eip = instrucciones.count + 1 ' Terminar ejecución
        Case Else
            Debug.Print "Instrucción no reconocida: " & partes(0)
    End Select
End Sub
Private Function ExtraerOperando(texto As String) As String
    If texto = "" Then
        ExtraerOperando = ""
        Exit Function
    End If
    
    ExtraerOperando = UCase(Trim(Replace(Replace(texto, ",", ""), " ", "")))
End Function
