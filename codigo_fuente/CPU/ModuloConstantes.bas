Attribute VB_Name = "ModuloConstantes"
' ========== CONSTANTES DE MEMORIA ==========
Public Const MEM_SIZE As Long = 65536
Public Const MEM_START As Long = 0
Public Const MEM_END As Long = 65535
Public Const STACK_START As Long = 64512
Public Const STACK_END As Long = 65535
' ========== CONSTANTES DE INSTRUCCIONES ==========
Public Const MAX_INSTRUCTION_LENGTH As Integer = 15
Public Const MAX_OPERANDS As Integer = 3

' ========== CONSTANTES DE PIPELINE ==========
Public Const PIPELINE_STAGES As Integer = 5
Public Const STAGE_FETCH As Integer = 0
Public Const STAGE_DECODE As Integer = 1
Public Const STAGE_EXECUTE As Integer = 2
Public Const STAGE_MEMORY As Integer = 3
Public Const STAGE_WRITEBACK As Integer = 4

' ========== CONSTANTES DE CACH? ==========
Public Const CACHE_SIZE As Long = 1024
Public Const CACHE_LINE_SIZE As Long = 64
Public Const CACHE_LINES As Long = 16

' ========== ENUMERACIONES ==========
Public Enum PIPELINE_STAGE
    Fetch = 0
    Decode = 1
    Execute = 2
    MEMORY_ACCESS = 3
    WRITE_BACK = 4
End Enum

Public Enum CACHE_POLICY
    LRU = 0
    FIFO = 1
    RANDOM = 2
End Enum

Public Enum INSTRUCTION_TYPE
    ARITHMETIC = 0
    LOGICAL = 1
    DATA_TRANSFER = 2
    CONTROL_FLOW = 3
    STACK = 4
End Enum

' ========== CONSTANTES DE VISUALIZACI?N ==========
Public Const DISPLAY_ROWS As Integer = 25
Public Const DISPLAY_COLS As Integer = 80
Public Const MEMORY_VIEW_ROWS As Integer = 16
Public Const MEMORY_VIEW_COLS As Integer = 8

