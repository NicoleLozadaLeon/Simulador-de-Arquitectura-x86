Attribute VB_Name = "Module1"
Option Explicit
Sub TraducirDesdeHoja()
    Dim codigoC As String
    codigoC = HojaEnsamblador.ObtenerCodigoC
    
    If codigoC = "" Then
        MsgBox "Por favor, ingresa un código C válido.", vbExclamation
        Exit Sub
    End If
    
    Dim ensamblador As String
    ensamblador = SimularTraduccion(codigoC)
    
    HojaEnsamblador.MostrarEnsamblador ensamblador
    HojaEnsamblador.ReiniciarSimulacion
    
    Dim lineas() As String
    lineas = Split(ensamblador, vbCrLf)
    HojaEnsamblador.AgregarInstrucciones lineas
End Sub
Sub EjecutarPasoDesdeHoja()
    If Not HojaEnsamblador.HayMasInstrucciones() Then
        If HojaEnsamblador.ObtenerCodigoC = "" Then
            MsgBox "Primero traduce el código C.", vbInformation
        Else
            MsgBox "Ejecución finalizada.", vbInformation
            HojaEnsamblador.MostrarInstruccion "(Fin)"
        End If
        Exit Sub
    End If
    
    Dim lineaInstruccion As String
    lineaInstruccion = HojaEnsamblador.ObtenerSiguienteInstruccion()
    HojaEnsamblador.MostrarInstruccion lineaInstruccion
    HojaEnsamblador.EjecutarInstruccion lineaInstruccion
    HojaEnsamblador.ActualizarRegistrosEnHoja
End Sub
Function SimularTraduccion(codigoC As String) As String
    Dim lineas() As String
    lineas = Split(codigoC, ";")
    
    Dim resultado As String
    Dim varA As Long, varB As Long, varC As Long
    
    Dim i As Long
    For i = 0 To UBound(lineas)
        Dim linea As String
        linea = Trim(lineas(i))
        If linea = "" Then GoTo Siguiente
        
        If InStr(linea, "=") > 0 Then
            Dim partes() As String
            partes = Split(linea, "=")
            Dim var As String: var = Trim(partes(0))
            Dim expr As String: expr = Trim(partes(1))
            
            If IsNumeric(expr) Then
                Select Case var
                    Case "a": varA = CLng(expr)
                    Case "b": varB = CLng(expr)
                    Case "c": varC = CLng(expr)
                End Select
                resultado = resultado & "mov " & MapearVar(var) & ", " & expr & vbCrLf
            ElseIf InStr(expr, "+") > 0 Then
                Dim sumandos() As String
                sumandos = Split(expr, "+")
                Dim op1 As String: op1 = Trim(sumandos(0))
                Dim op2 As String: op2 = Trim(sumandos(1))
                
                resultado = resultado & "mov eax, " & MapearVar(op1) & vbCrLf
                If IsNumeric(op2) Then
                    resultado = resultado & "add eax, " & op2 & vbCrLf
                Else
                    resultado = resultado & "add eax, " & MapearVar(op2) & vbCrLf
                End If
                resultado = resultado & "mov " & MapearVar(var) & ", eax" & vbCrLf
            ElseIf InStr(expr, "-") > 0 Then
                Dim restas() As String
                restas = Split(expr, "-")
                Dim minuendo As String: minuendo = Trim(restas(0))
                Dim sustraendo As String: sustraendo = Trim(restas(1))
                
                resultado = resultado & "mov eax, " & MapearVar(minuendo) & vbCrLf
                If IsNumeric(sustraendo) Then
                    resultado = resultado & "sub eax, " & sustraendo & vbCrLf
                Else
                    resultado = resultado & "sub eax, " & MapearVar(sustraendo) & vbCrLf
                End If
                resultado = resultado & "mov " & MapearVar(var) & ", eax" & vbCrLf
            End If
        End If
Siguiente:
    Next i
    
    If resultado = "" Then resultado = "; No se pudo traducir." & vbCrLf & "; Usa: a = 5; b = a + 3;"
    SimularTraduccion = resultado
End Function
Function MapearVar(nombreVar As String) As String
    Select Case nombreVar
        Case "a": MapearVar = "dword [a]"
        Case "b": MapearVar = "dword [b]"
        Case "c": MapearVar = "dword [c]"
        Case Else: MapearVar = nombreVar
    End Select
End Function
