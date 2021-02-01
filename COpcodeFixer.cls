VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COpcodeFixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // COpcodeFixer.cls - opcodes handlers module loader
' // by The trick 2021
' //

Option Explicit

Private Const MODULE_NAME       As String = "COpcodeFixer"

Private Const MODULE_SIGNATURE  As Long = &H7866706F
Private Const MODULE_VERSION    As Long = 1

Private Type tModuleHeader
    lModuleSignature    As Long
    lModuleVersion      As Long
End Type

Private Type tOpcodeHandler
    lOpcodeID           As Long
    lHandlerOffset      As Long
End Type

Private Type tImport
    lOffset             As Long
    bResolved           As Boolean
    sName               As String
End Type

Private m_tImports()    As tImport          ' // Import data
Private m_tOpcodes()    As tOpcodeHandler   ' // New handlers
Private m_pModuleCode   As PTR              ' // Code in RWX-memory
Private m_bPrepared     As Boolean

' // Initialize module from byte array raw data
Public Sub Initialize( _
           ByRef bModule() As Byte)
    Const PROC_NAME = "Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim cStream     As CROStream
    Dim tHeader     As tModuleHeader
    Dim tOpcodes()  As tOpcodeHandler
    Dim tImports()  As tImport
    Dim lRelocs()   As Long
    Dim lCodeSize   As Long
    Dim pCode       As PTR
    Dim lIndex      As Long
    Dim pRelEntry   As PTR
    
    On Error GoTo err_handler
    
    Set cStream = New CROStream
    
    cStream.InitializeFromByteArray bModule()
    
    ' // Read header
    cStream.ReadData VarPtr(tHeader), LenB(tHeader)
    
    If tHeader.lModuleSignature <> MODULE_SIGNATURE Then
        Err.Raise 5, , "Invalid signature"
    ElseIf tHeader.lModuleVersion > MODULE_VERSION Then
        Err.Raise 5, , "Unsupported module version"
    End If
    
    ' // Read defined opcodes
    tOpcodes = ExtractOpcodes(cStream)
    
    ' // Read imports
    cStream.Align 4
    tImports = ExtractImports(cStream)
    
    ' // Skip code
    lCodeSize = cStream.ReadLong()
    cStream.Position = cStream.Position + lCodeSize
    
    ' // Read relocations
    lRelocs = ExtractRelocations(cStream)
    
    ' // Allocate memory for code
    pCode = VirtualAlloc(0, cStream.Size, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pCode = NULL_PTR Then
        Err.Raise 7
    End If
    
    ' // Copy code
    cStream.Position = 0
    cStream.ReadData pCode, cStream.Size
    
    ' // Apply relocations
    For lIndex = 0 To SafeArrayElementsCount(ArrPtr(lRelocs)) - 1
        
        GetMemPtr ByVal pCode + lRelocs(lIndex), pRelEntry
        pRelEntry = pRelEntry + pCode
        PutMemPtr ByVal pCode + lRelocs(lIndex), pRelEntry
        
    Next
    
    CleanAll
    
    m_pModuleCode = pCode
    m_tOpcodes = tOpcodes
    m_tImports = tImports
    
    Exit Sub
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Sub

Public Property Get NumberOfOpHandlers() As Long
    NumberOfOpHandlers = SafeArrayElementsCount(ArrPtr(m_tOpcodes))
End Property

' // Get opcode by index
Public Property Get OpcodeID( _
                    ByVal lIndex As Long) As Long
    Const PROC_NAME = "OpcodeID_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If lIndex < 0 Or lIndex >= SafeArrayElementsCount(ArrPtr(m_tOpcodes)) Then
        Err.Raise 9, FULL_PROC_NAME
    End If
    
    OpcodeID = m_tOpcodes(lIndex).lOpcodeID
    
End Property

' // Get opcode handler by index
Public Property Get Handler( _
                    ByVal lIndex As Long) As PTR
    Const PROC_NAME = "Handler_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If Not m_bPrepared Then
        Err.Raise 5, FULL_PROC_NAME, "Module isn't prepared"
    ElseIf lIndex < 0 Or lIndex >= SafeArrayElementsCount(ArrPtr(m_tOpcodes)) Then
        Err.Raise 9, FULL_PROC_NAME
    End If
    
    Handler = m_tOpcodes(lIndex).lHandlerOffset + m_pModuleCode
    
End Property

' // Check if there is unresolved import items
' // If module has unresolved import it can't be used
Public Property Get HasUnresolvedImport() As Boolean
    Dim lIndex  As Long
    
    For lIndex = 0 To SafeArrayElementsCount(ArrPtr(m_tImports)) - 1
        If Not m_tImports(lIndex).bResolved Then
            HasUnresolvedImport = True
            Exit Property
        End If
    Next
    
End Property

' // Resolve import
' // It fills the import table with the addresses from IImportProvider
Public Sub ResolveImport( _
           ByVal cImportProvider As IImportProvider)
    Const PROC_NAME = "ResolveImport", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lIndex      As Long
    Dim pAddr       As PTR
    Dim bPrepared   As Boolean
    
    On Error GoTo err_handler
    
    bPrepared = True
    
    For lIndex = 0 To SafeArrayElementsCount(ArrPtr(m_tImports)) - 1
        With m_tImports(lIndex)
            If cImportProvider.Exists(.sName) Then
                
                pAddr = cImportProvider.Item(.sName)
                
                PutMem4 ByVal m_pModuleCode + .lOffset, pAddr
                                
                .bResolved = True
            
            Else
                bPrepared = False
            End If
        End With
    Next
          
    m_bPrepared = bPrepared
    
    Exit Sub
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
           
End Sub

' // Clean up
Private Sub CleanAll()
    
    If m_pModuleCode Then
        VirtualFree m_pModuleCode, 0, MEM_RELEASE
    End If
    
    Erase m_tOpcodes
    Erase m_tImports
    
    m_bPrepared = False
    
End Sub

' // Extract defined opcodes from stream
Private Function ExtractOpcodes( _
                 ByVal cStream As CROStream) As tOpcodeHandler()
    Const PROC_NAME = "ExtractOpcodes", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lCount  As Long
    Dim lIndex  As Long
    Dim tOut()  As tOpcodeHandler
    
    On Error GoTo err_handler
    
    lCount = cStream.ReadLong()
    
    If lCount > 0 Then
        
        ReDim tOut(lCount - 1)
        
        For lIndex = 0 To lCount - 1
            
            With tOut(lIndex)
            
                .lOpcodeID = cStream.ReadByte()
                
                If .lOpcodeID >= &HFB Then
                    .lOpcodeID = .lOpcodeID * &H100 Or cStream.ReadByte()
                End If
                
                .lHandlerOffset = cStream.ReadLong()
                
            End With
            
        Next
        
    End If
    
    ExtractOpcodes = tOut
    
    Exit Function
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Function

' // Extract needed imports from stream
Private Function ExtractImports( _
                 ByVal cStream As CROStream) As tImport()
    Const PROC_NAME = "ExtractImports", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lCount  As Long
    Dim lIndex  As Long
    Dim tOut()  As tImport
    
    On Error GoTo err_handler
    
    lCount = cStream.ReadLong()
    
    If lCount > 0 Then
        
        ReDim tOut(lCount - 1)
        
        For lIndex = 0 To lCount - 1
            
            tOut(lIndex).lOffset = cStream.Position
            cStream.Position = cStream.Position + SIZEOF_PTR
            
        Next
        
        For lIndex = 0 To lCount - 1
        
            tOut(lIndex).sName = cStream.ReadMBZStringToUTF16()
            tOut(lIndex).bResolved = False
            
        Next
        
    End If
    
    ExtractImports = tOut
    
    Exit Function
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Function

' // Extract relocations from stream
Private Function ExtractRelocations( _
                 ByVal cStream As CROStream) As Long()
    Const PROC_NAME = "ExtractRelocations", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lCount  As Long
    Dim lOut()  As Long
    
    On Error GoTo err_handler
    
    lCount = cStream.ReadLong()
    
    If lCount > 0 Then
        
        ReDim lOut(lCount - 1)
        
        cStream.ReadData VarPtr(lOut(0)), lCount * Len(lOut(0))
        
    End If
    
    ExtractRelocations = lOut
    
    Exit Function
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Function

Private Sub Class_Terminate()
    CleanAll
End Sub
