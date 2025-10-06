Attribute VB_Name = "ModuloRegistros"
' ===== ModuloRegistros =====
Option Explicit

' Registros de propósito general
Public EAX As Long, EBX As Long, ECX As Long, EDX As Long
Public ESI As Long, EDI As Long, EBP As Long, ESP As Long
Public eip As Long

' Flags
Public ZF As Boolean, SF As Boolean, CF As Boolean, OF As Boolean
Public PF As Boolean, AF As Boolean

' Estado
Public SimulacionEnCurso As Boolean
Public instrucciones As Collection
Public Sub InicializarRegistros()
    EAX = 0:
    EBX = 0:
    ECX = 0:
    EDX = 0
    ESI = 0:
    EDI = 0:
    EBP = 0:
    ESP = STACK_START
    eip = 0
    ZF = False: SF = False: CF = False: OF = False: PF = False: AF = False
End Sub
