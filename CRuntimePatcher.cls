VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRuntimePatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CRuntimePatcher.cls - VM runtime patcher
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME       As String = "CRuntimePatcher"

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

Private Type tCdeclDeclarePatchData
    pPatchPos       As PTR      ' // Where patch
    pPatchTo        As PTR      ' // Destination address
    lOriginalOffset As Long     ' // Original offset (to restore)
End Type

Private m_cVba6ModInfo      As CPEModuleInfo
Private m_pBugTable         As PTR
Private m_pCodeTables()     As PTR
Private m_pImports()        As PTR
Private m_cNewOpcodes       As COpcodeFixer
Private m_bCDeclEnabled     As Boolean
Private m_bIsInitialized    As Boolean
Private m_pOriginHandlers() As PTR
Private m_tCdeclDeclPatch   As tCdeclDeclarePatchData

Implements IImportProvider

' // Initialize module
Public Sub Initialize()
    Const PROC_NAME = "Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim cScanner        As CSignaturesScanner
    Dim pBugTable       As PTR
    Dim pPCodeTables()  As PTR
    Dim pImports()      As PTR
    Dim tDeclPatch      As tCdeclDeclarePatchData
    
    On Error GoTo err_handler
    
    If m_bIsInitialized Then
        Exit Sub
    End If
    
    Set cScanner = New CSignaturesScanner
    
    cScanner.InitializeFromByteArray LoadResData(102, RT_RCDATA)
    
    pBugTable = ExtractBugTable(cScanner)
    
    If pBugTable = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME, "Unable to find bug table"
    End If
    
    pPCodeTables = ExtractPCodesTables(cScanner)
    
    If SafeArrayElementsCount(ArrPtr(pPCodeTables)) <= 0 Then
        Err.Raise 5, FULL_PROC_NAME, "Unable to find P-Codes tables"
    End If
    
    pImports = ExtractImport(cScanner, pPCodeTables)
    
    If SafeArrayElementsCount(ArrPtr(pImports)) <= 0 Then
        Err.Raise 5, FULL_PROC_NAME, "Unable to find import"
    End If
    
    tDeclPatch = ExtractCdeclDeclarePatch(cScanner)
    
    If tDeclPatch.pPatchPos = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME, "Unable to find Cdecl declare code"
    End If
    
    m_pBugTable = pBugTable
    m_pCodeTables = pPCodeTables
    m_pImports = pImports
    m_tCdeclDeclPatch = tDeclPatch
    
    m_bIsInitialized = True
    
    Exit Sub

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Sub

Public Property Get CDeclEnabled() As Boolean
    CDeclEnabled = m_bCDeclEnabled
End Property

' // Enable/disable CDECL support
Public Property Let CDeclEnabled( _
                    ByVal bValue As Boolean)
                        
    If m_bCDeclEnabled = bValue Then Exit Property
    
    If bValue Then
        EnableCdecl
    Else
        DisableCdecl
    End If
                        
End Property

' // Disable CDECL
Private Sub DisableCdecl()
    Const PROC_NAME = "DisableCdecl", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    Dim lOpcodeIndex    As Long
    
    On Error GoTo err_handler
    
    ' // Patch bug table
    PatchBugTable False
    
    ' // Replace stubs addresses
    For lOpcodeIndex = 0 To m_cNewOpcodes.NumberOfOpHandlers - 1
        ReplaceOpcodeHandler m_cNewOpcodes.OpcodeID(lOpcodeIndex), m_pOriginHandlers(lOpcodeIndex)
    Next
    
    PatchDeclareCdeclJmp False
    
    m_bCDeclEnabled = False
    
    Exit Sub

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Sub

' // Enable CDECL
Private Sub EnableCdecl()
    Const PROC_NAME = "EnableCdecl", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    Dim lOpcodeIndex    As Long
    
    On Error GoTo err_handler
    
    ' // Patch bug table
    ' // Now the crash won't occur but compiler will compile p-code with the stubs
    ' // which raises intrnal error exception. So we need to replace those stubs
    PatchBugTable True
    
    If m_cNewOpcodes Is Nothing Then
        InitializeNewOpcodes
    End If
    
    ' // Replace stubs addresses
    For lOpcodeIndex = 0 To m_cNewOpcodes.NumberOfOpHandlers - 1
        ReplaceOpcodeHandler m_cNewOpcodes.OpcodeID(lOpcodeIndex), m_cNewOpcodes.Handler(lOpcodeIndex)
    Next
    
    PatchDeclareCdeclJmp True
    
    m_bCDeclEnabled = True
    
    Exit Sub

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
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

Private Function OpcodeTableEntryFromOpcodeId( _
                 ByVal lOpcodeID As Long) As PTR
    Const PROC_NAME = "OpcodeTableEntryFromOpcodeId", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If lOpcodeID <= &HFA Then
        
        If lOpcodeID < 0 Then
            Err.Raise 5, PROC_NAME
        End If
        
        OpcodeTableEntryFromOpcodeId = m_pCodeTables(0) + lOpcodeID * SIZEOF_PTR
        
    Else
    
        Select Case lOpcodeID \ &H100
        Case &HFB
            OpcodeTableEntryFromOpcodeId = m_pCodeTables(1) + (lOpcodeID And &HFF) * SIZEOF_PTR
        Case &HFC
            OpcodeTableEntryFromOpcodeId = m_pCodeTables(2) + (lOpcodeID And &HFF) * SIZEOF_PTR
        Case &HFD
            OpcodeTableEntryFromOpcodeId = m_pCodeTables(3) + (lOpcodeID And &HFF) * SIZEOF_PTR
        Case &HFE
            OpcodeTableEntryFromOpcodeId = m_pCodeTables(4) + (lOpcodeID And &HFF) * SIZEOF_PTR
        Case &HFF
        
            If (lOpcodeID And &HFF) >= &H47 Then
                Err.Raise 5, PROC_NAME
            End If
            
            OpcodeTableEntryFromOpcodeId = m_pCodeTables(5) + (lOpcodeID And &HFF) * SIZEOF_PTR
            
        Case Else
            Err.Raise 5, PROC_NAME
        End Select
        
    End If
    
End Function

Private Sub InitializeNewOpcodes()
    Const PROC_NAME = "InitializeNewOpcodes", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim cNewOpcodes As COpcodeFixer
    Dim lIndex      As Long
    
    On Error GoTo err_handler
    
    Set cNewOpcodes = New COpcodeFixer
    
    cNewOpcodes.Initialize LoadResData(101, RT_RCDATA)
    
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

Private Sub PatchDeclareCdeclJmp( _
            ByVal bPatch As Boolean)
    Const PROC_NAME = "PatchDeclareCdeclJmp", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lOldProtect As Long
    Dim lOffset     As Long
    
    If VirtualProtect(m_tCdeclDeclPatch.pPatchPos, &H4, PAGE_EXECUTE_READWRITE, lOldProtect) = 0 Then
        Err.Raise 7, PROC_NAME
    End If
    
    Debug.Print Hex$(m_tCdeclDeclPatch.pPatchPos)
    
    If bPatch Then
        lOffset = m_tCdeclDeclPatch.pPatchTo - (m_tCdeclDeclPatch.pPatchPos + 4)
    Else
        lOffset = m_tCdeclDeclPatch.lOriginalOffset
    End If
    
    PutMem4 ByVal m_tCdeclDeclPatch.pPatchPos, lOffset
    
    VirtualProtect m_tCdeclDeclPatch.pPatchPos, &H4, lOldProtect, lOldProtect
    
End Sub

Private Sub PatchBugTable( _
            ByVal bPatch As Boolean)
    Const PROC_NAME = "PatchBugTable", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lOldProtect As Long
    Dim bNewValue   As Byte
    
    If VirtualProtect(m_pBugTable, &H1E, PAGE_EXECUTE_READWRITE, lOldProtect) = 0 Then
        Err.Raise 7, PROC_NAME
    End If
    
    'Debug.Print Hex$(m_pBugTable)
    
    If bPatch Then
        bNewValue = 9
    Else
        bNewValue = &H19
    End If
    
    PutMem1 ByVal m_pBugTable + &H1D, bNewValue
    
    VirtualProtect m_pBugTable, &H1E, lOldProtect, lOldProtect
           
End Sub

Private Function ExtractCdeclDeclarePatch( _
                 ByVal cScanner As CSignaturesScanner) As tCdeclDeclarePatchData
    Const PROC_NAME = "ExtractCdeclDeclarePatch", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pTest       As PTR
    Dim pTarget     As PTR
    Dim pJmpTo      As PTR
    Dim lOffset     As Long
    Dim tSecRange   As IMAGE_SECTION_HEADER
    Dim iMovECXEDI  As Integer
    
    On Error GoTo err_handler
    
    ' // Check range
    tSecRange = m_cVba6ModInfo.SectionByName(".text")
    
    ' // Set specified sequence
    cScanner.StartSearch "Declare_Cdecl_Check"

    Do
        
        pTest = cScanner.SearchNext()
        
        If pTest = NULL_PTR Then
            Exit Function
        End If
        
        pTarget = cScanner.Target("CompileDeclare")
        
        If pTarget >= tSecRange.VirtualAddress + m_cVba6ModInfo.BaseAddress And _
            pTarget < tSecRange.VirtualAddress + tSecRange.VirtualSize + m_cVba6ModInfo.BaseAddress Then
            
            ' // Check original address
            pJmpTo = cScanner.Target("CdeclCompileErr31")
            
            GetMem4 ByVal pJmpTo, lOffset
            
            pTest = lOffset + (pJmpTo + 4)
            
            ' // Check mov ecx, edi
            GetMem2 ByVal pTarget, iMovECXEDI
            
            ' // Found
            If iMovECXEDI = &HCF8B And pTest <> pTarget Then
                
                ExtractCdeclDeclarePatch.pPatchTo = pTarget
                ExtractCdeclDeclarePatch.pPatchPos = pJmpTo
                
                GetMem4 ByVal ExtractCdeclDeclarePatch.pPatchPos, ExtractCdeclDeclarePatch.lOriginalOffset
                
                Exit Function
                
            End If
            
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


' // Extract bug table ptr
Private Function ExtractBugTable( _
                 ByVal cScanner As CSignaturesScanner) As PTR
    Const PROC_NAME = "ExtractBugTable", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pTable      As PTR
    Dim tSecRange   As IMAGE_SECTION_HEADER
    Dim bBugValue   As Byte
    
    On Error GoTo err_handler
    
    cScanner.StartSearch "Bug_Signature"
    
    tSecRange = m_cVba6ModInfo.SectionByName(".text")
    
    Do
        
        ' // Search for code sequence where table is referenced
        If cScanner.SearchNext() = NULL_PTR Then
            Exit Function
        End If
        
        ' // Get pointer to address to table
        pTable = cScanner.Target("BugTablePtr")

        ' // Check if address is within .text section
        If pTable >= tSecRange.VirtualAddress + m_cVba6ModInfo.BaseAddress Or _
            pTable + &H1D < tSecRange.VirtualAddress + tSecRange.VirtualSize + m_cVba6ModInfo.BaseAddress Then
            
            ' // Check value
            GetMem1 ByVal pTable + &H1D, bBugValue
            
            If bBugValue = &H19 Then
                
                ExtractBugTable = pTable
                Exit Function
                
            End If
            
        End If
        
    Loop While True
    
    Exit Function

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Function

Private Sub Class_Initialize()
    
    Set m_cVba6ModInfo = New CPEModuleInfo
    
    m_cVba6ModInfo.LoadAndLock "vba6"
    
End Sub

Private Sub Class_Terminate()
    
    If m_bCDeclEnabled Then
        DisableCdecl
    End If
    
End Sub

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
    Case "__imp_HrDefSetIndex":         IImportProvider_Item = m_pImports(IMP_HrDefSetIndex)
    Case "__imp_VBAEventImportCall":    IImportProvider_Item = m_pImports(IMP_VBAEventImportCall)
    Case "__imp_VBAEventProcExit":      IImportProvider_Item = m_pImports(IMP_VBAEventProcExit)
    Case "_g_pEventMonitorsEnabled":    IImportProvider_Item = m_pImports(IMP_g_pEventMonitorsEnabled)
    Case "__imp_EbRaiseExceptionCode":  IImportProvider_Item = m_pImports(IMP_EbRaiseExceptionCode)
    Case "__imp_AllocStackUnk":         IImportProvider_Item = m_pImports(IMP_AllocStackUnk)
    Case "__imp_HresultCheck":          IImportProvider_Item = m_pImports(IMP_HresultCheck)
    Case "_g_ExceptFlags":              IImportProvider_Item = m_pImports(IMP_g_ExceptFlags)
    Case "_g_DispTable":                IImportProvider_Item = m_pImports(IMP_g_DispTable)
    Case Else
        Err.Raise 5, PROC_NAME
    End Select
    
    Exit Property

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Property


