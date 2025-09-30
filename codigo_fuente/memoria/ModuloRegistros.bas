Attribute VB_Name = "ModuloRegistros"
' ModuloRegistros.bas
Option Explicit

' Registros de propósito general (32 bits)
Public EAX As Long
Public EBX As Long
Public ECX As Long
Public EDX As Long
' Registros de segmento
Public CS As Integer    ' Code Segment
Public DS As Integer    ' Data Segment
Public SS As Integer    ' Stack Segment
Public ES As Integer    ' Extra Segment

' Registros de puntero
Public EIP As Long      ' Instruction Pointer
Public ESP As Long      ' Stack Pointer
Public EBP As Long      ' Base Pointer

' Registros de índice
Public ESI As Long      ' Source Index
Public EDI As Long      ' Destination Index

' Flags
Public ZF As Boolean    ' Zero Flag
Public SF As Boolean    ' Sign Flag
Public CF As Boolean    ' Carry Flag
Public OF As Boolean    ' Overflow Flag
Public PF As Boolean    ' Parity Flag
Public AF As Boolean    ' Auxiliary Flag

' Inicializar todos los registros
Public Sub InicializarRegistros()
    EAX = 0
    EBX = 0
    ECX = 0
    EDX = 0
    CS = &H1000
    DS = &H2000
    SS = &H3000
    ES = &H4000
    EIP = &H0
    ESP = &HFFFF
    EBP = 0
    ESI = 0
    EDI = 0
    ZF = False
    SF = False
    CF = False
    OF = False
    PF = False
    AF = False
End Sub
