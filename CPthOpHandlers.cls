VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPthOpHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CPthOpHandlers.cls - serves patches to replace unimplemented opcodes
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME   As String = "CPthOpHandlers"

Private Enum eOpcodes
    
    OP_ImpAdCallHresult = 9
    
End Enum

Private Enum eImport
    
    IMP_HrDefSetIndex
    IMP_VBAEventImportCall
    IMP_VBAEventProcExit
    IMP_g_pEventMonitorsEnabled
    IMP_EbRaiseExceptionCode
    IMP_AllocStackUnk
    IMP_HresultCheck
    IMP_g_ExceptFlags
    IMP_g_DispTable
    
    IMP_TOTAL
    
End Enum

Implements IPatchEntry
Implements IImportProvider

Private m_cVba6ModInfo      As CPEModuleInfo
Private m_pHandlers()       As PTR
Private m_pOriginHandlers() As PTR
Private m_pImport()         As PTR
Private m_bPatched          As Boolean
Private m_cNewOpcodes       As COpcodeFixer
Private m_bIsInitialized    As Boolean

Private Sub Class_Initialize()

    Set m_cVba6ModInfo = New CPEModuleInfo
    
    m_cVba6ModInfo.LoadAndLock "vba6"

End Sub

Private Sub Class_Terminate()
    
    If m_bPatched Then
        
        Patch_internal False
        m_bPatched = False
        
    End If
    
    m_bIsInitialized = False
    
    Set m_cNewOpcodes = Nothing
    
End Sub

' //
' // IImportProvider implementation
' //

Private Property Get IImportProvider_Exists( _
                     ByRef sName As String) As Boolean
    
    Select Case sName
    Case "__imp_HrDefSetIndex", "__imp_VBAEventImportCall", "__imp_VBAEventProcExit", _
         "_g_pEventMonitorsEnabled", "__imp_EbRaiseExceptionCode", "__imp_AllocStackUnk", _
         "__imp_HresultCheck", "_g_ExceptFlags", "_g_DispTable"
        IImportProvider_Exists = True
    End Select
    
End Property

Private Property Get IImportProvider_Item( _
                     ByRef sName As String) As PTR
    Const PROC_NAME = "IImportProvider_Item", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Select Case sName
    Case "__imp_HrDefSetIndex":         IImportProvider_Item = m_pImport(IMP_HrDefSetIndex)
    Case "__imp_VBAEventImportCall":    IImportProvider_Item = m_pImport(IMP_VBAEventImportCall)
    Case "__imp_VBAEventProcExit":      IImportProvider_Item = m_pImport(IMP_VBAEventProcExit)
    Case "_g_pEventMonitorsEnabled":    IImportProvider_Item = m_pImport(IMP_g_pEventMonitorsEnabled)
    Case "__imp_EbRaiseExceptionCode":  IImportProvider_Item = m_pImport(IMP_EbRaiseExceptionCode)
    Case "__imp_AllocStackUnk":         IImportProvider_Item = m_pImport(IMP_AllocStackUnk)
    Case "__imp_HresultCheck":          IImportProvider_Item = m_pImport(IMP_HresultCheck)
    Case "_g_ExceptFlags":              IImportProvider_Item = m_pImport(IMP_g_ExceptFlags)
    Case "_g_DispTable":                IImportProvider_Item = m_pImport(IMP_g_DispTable)
    Case Else
        Err.Raise 5, PROC_NAME
    End Select
    
    Exit Property

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Property

' //
' // IPatchEntry implementation
' //

Private Sub IPatchEntry_Initialize( _
            ByVal cScanner As CSignaturesScanner)
    Const PROC_NAME = "IPatchEntry_Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pHandlers() As PTR
    Dim pImport()   As PTR
    
    On Error GoTo err_handler
    
    pHandlers = ExtractPCodesTables(cScanner)
    pImport = ExtractImport(cScanner, pHandlers())
    
    m_pHandlers = pHandlers
    m_pImport = pImport
    
    InitializeNewOpcodes
    
    m_bIsInitialized = True
    
    Exit Sub

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Sub

' // Get opcode handler address from opcode id
Private Function OpcodeTableEntryFromOpcodeId( _
                 ByVal lOpcodeID As Long) As PTR
    Const PROC_NAME = "OpcodeTableEntryFromOpcodeId", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If lOpcodeID <= &HFA Then
        
        If lOpcodeID < 0 Then
            Err.Raise 5, PROC_NAME
        End If
        
        OpcodeTableEntryFromOpcodeId = m_pHandlers(0) + lOpcodeID * SIZEOF_PTR
        
    Else
    
        Select Case lOpcodeID \ &H100
        Case &HFB
            OpcodeTableEntryFromOpcodeId = m_pHandlers(1) + (lOpcodeID And &HFF) * SIZEOF_PTR
        Case &HFC
            OpcodeTableEntryFromOpcodeId = m_pHandlers(2) + (lOpcodeID And &HFF) * SIZEOF_PTR
        Case &HFD
            OpcodeTableEntryFromOpcodeId = m_pHandlers(3) + (lOpcodeID And &HFF) * SIZEOF_PTR
        Case &HFE
            OpcodeTableEntryFromOpcodeId = m_pHandlers(4) + (lOpcodeID And &HFF) * SIZEOF_PTR
        Case &HFF
        
            If (lOpcodeID And &HFF) >= &H47 Then
                Err.Raise 5, PROC_NAME
            End If
            
            OpcodeTableEntryFromOpcodeId = m_pHandlers(5) + (lOpcodeID And &HFF) * SIZEOF_PTR
            
        Case Else
            Err.Raise 5, PROC_NAME
        End Select
        
    End If
    
End Function

' // Initialize opcodes module
Private Sub InitializeNewOpcodes()
    Const PROC_NAME = "InitializeNewOpcodes", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim cNewOpcodes As COpcodeFixer
    Dim lIndex      As Long
    
    On Error GoTo err_handler
    
    Set cNewOpcodes = New COpcodeFixer
    
    ' // Load opcodes module
    cNewOpcodes.Initialize LoadResData(RID_OPCODES_HANDLERS, RT_RCDATA)
    
    cNewOpcodes.ResolveImport Me
    
    If cNewOpcodes.HasUnresolvedImport Then
        Err.Raise 5, PROC_NAME, "Has unresolved import"
    End If
    
    If cNewOpcodes.NumberOfOpHandlers > 0 Then
        ReDim m_pOriginHandlers(cNewOpcodes.NumberOfOpHandlers - 1)
    Else
        Erase m_pOriginHandlers
    End If
    
    ' // Save original handlers
    For lIndex = 0 To cNewOpcodes.NumberOfOpHandlers - 1
        GetMemPtr ByVal OpcodeTableEntryFromOpcodeId(cNewOpcodes.OpcodeID(lIndex)), m_pOriginHandlers(lIndex)
    Next
    
    Set m_cNewOpcodes = cNewOpcodes
    
    Exit Sub

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Sub

' // Extract P-code handlers tables
Private Function ExtractPCodesTables( _
                 ByVal cScanner As CSignaturesScanner) As PTR()
    Const PROC_NAME = "ExtractPCodesTable", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pTables()   As PTR
    Dim tSecRange   As IMAGE_SECTION_HEADER
    Dim lIndex      As Long
    Dim lHdlrIndex  As Long
    Dim lHdlrCount  As Long
    Dim pHandler    As PTR
    
    On Error GoTo err_handler
    
    ReDim pTables(5)
    
    ' // Scan specified sequence
    cScanner.StartSearch "VM_DispTables"
    
    ' // All the handlers are within .ENGINE section
    tSecRange = m_cVba6ModInfo.SectionByName("ENGINE")
    
    Do
    
        If cScanner.SearchNext() = NULL_PTR Then
            Exit Function
        End If

        For lIndex = 0 To UBound(pTables)
        
            pTables(lIndex) = cScanner.Target(lIndex)
            
            If lHdlrIndex = 5 Then
                lHdlrCount = &H47
            Else
                lHdlrCount = &H100
            End If
            
            If pTables(lIndex) < tSecRange.VirtualAddress + m_cVba6ModInfo.BaseAddress Or _
                pTables(lIndex) + (lHdlrCount * SIZEOF_PTR) > tSecRange.VirtualAddress + _
                tSecRange.VirtualSize + m_cVba6ModInfo.BaseAddress Then
                Exit For
            End If
            
            For lHdlrIndex = 0 To lHdlrCount - 1
                
                GetMemPtr ByVal pTables(lIndex) + lHdlrIndex * SIZEOF_PTR, pHandler
                
                If pHandler < tSecRange.VirtualAddress + m_cVba6ModInfo.BaseAddress Or _
                    pHandler >= tSecRange.VirtualAddress + tSecRange.VirtualSize + m_cVba6ModInfo.BaseAddress Then
                    Exit For
                End If
                
            Next
            
        Next
        
        If lIndex > UBound(pTables) Then
            
            ExtractPCodesTables = pTables
            Exit Function
            
        End If
        
    Loop While True
    
    Exit Function

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Function

' // Extract unxported functions / global variables
Private Function ExtractImport( _
                 ByVal cScanner As CSignaturesScanner, _
                 ByRef pPCodeTables() As PTR) As PTR()
    Const PROC_NAME = "ExtractImport", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pImport()           As PTR
    Dim pImpAdCallHresult   As PTR
    Dim pTest               As PTR
    Dim lOffset             As Long
    Dim pEbRaiseEx          As PTR
    
    On Error GoTo err_handler
    
    ReDim pImport(IMP_TOTAL - 1)
    
    ' // Get ImpAdCallHresult handler
    GetMemPtr ByVal pPCodeTables(0) + OP_ImpAdCallHresult * SIZEOF_PTR, pImpAdCallHresult
    
    ' // Set specified sequence
    cScanner.StartSearch "ImpAdCallHresult"

    Do
        
        pTest = cScanner.SearchNext()
        
        If pTest = NULL_PTR Then
            Exit Function
        End If
        
        If pTest = pImpAdCallHresult Then
            
            pImport(IMP_HrDefSetIndex) = cScanner.Target("HrDefSetIndexCall")
            pImport(IMP_VBAEventImportCall) = cScanner.Target("VbaEventImportCall")
            pImport(IMP_VBAEventProcExit) = cScanner.Target("VbaEventProcExitCall")
            pImport(IMP_g_pEventMonitorsEnabled) = cScanner.Target("EventMonitorsEnabledPtr")
            pImport(IMP_AllocStackUnk) = cScanner.Target("AllocStackUnkCall")
            pImport(IMP_HresultCheck) = cScanner.Target("HresultCheckPtr")
            pImport(IMP_g_ExceptFlags) = cScanner.Target("ExceptFlagPtr")
            
            pEbRaiseEx = GetProcAddress(m_cVba6ModInfo.BaseAddress, "EbRaiseExceptionCode")
            
            If pEbRaiseEx = NULL_PTR Then
                Exit Function
            End If
            
            pImport(IMP_EbRaiseExceptionCode) = pEbRaiseEx
            pImport(IMP_g_DispTable) = pPCodeTables(0)
            
            ExtractImport = pImport
            
            Exit Function
            
        End If
        
    Loop While True
    
    Exit Function

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Function


Private Sub IPatchEntry_Patch()
    Const PROC_NAME = "IPatchEntry_Patch", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
    
    If Not m_bIsInitialized Then
        Err.Raise 5, , "Not initialized"
    End If
    
    Patch_internal True
    m_bPatched = True
    
    Exit Sub
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
         
End Sub

Private Sub IPatchEntry_Unpatch()
    Const PROC_NAME = "IPatchEntry_Unpatch", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
    
    Patch_internal False
    m_bPatched = False
    
    Exit Sub
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
         
End Sub


Private Sub Patch_internal( _
            ByVal bPatch As Boolean)
    Const PROC_NAME = "Patch_internal", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lOpcodeIndex    As Long
    
    If bPatch Then
    
        ' // Replace stubs addresses
        For lOpcodeIndex = 0 To m_cNewOpcodes.NumberOfOpHandlers - 1
            ReplaceOpcodeHandler m_cNewOpcodes.OpcodeID(lOpcodeIndex), m_cNewOpcodes.Handler(lOpcodeIndex)
        Next

    Else
    
        ' // Replace stubs addresses
        For lOpcodeIndex = 0 To m_cNewOpcodes.NumberOfOpHandlers - 1
            ReplaceOpcodeHandler m_cNewOpcodes.OpcodeID(lOpcodeIndex), m_pOriginHandlers(lOpcodeIndex)
        Next
    
    End If
    
End Sub

' // Replace opcode handler
' // Returns previous handler
Private Function ReplaceOpcodeHandler( _
                 ByVal lOpcode As Long, _
                 ByVal pNewHandler As PTR) As PTR
    Const PROC_NAME = "ReplaceOpcodeHandler", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lOldProtect As Long
    Dim pTableEntry As PTR
    
    On Error GoTo err_handler
    
    pTableEntry = OpcodeTableEntryFromOpcodeId(lOpcode)
    
    If VirtualProtect(pTableEntry, SIZEOF_PTR, PAGE_EXECUTE_READWRITE, lOldProtect) = 0 Then
        Err.Raise 7, PROC_NAME
    End If
    
    GetMemPtr ByVal pTableEntry, ReplaceOpcodeHandler
    PutMemPtr ByVal pTableEntry, pNewHandler
    
    VirtualProtect pTableEntry, SIZEOF_PTR, lOldProtect, lOldProtect
     
    Exit Function

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
     
End Function






