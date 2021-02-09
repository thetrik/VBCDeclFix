VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPthUserCDeclFn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CPthUserCDeclFn.cls - serves patch which enable CDecl keyword in user functions
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME   As String = "CPthUserCDeclFn"

Implements IPatchEntry
Implements IImportProvider

' // This structure hold data that ensures passing CDecl keyword after procedure name
Private Type tPCRCheckId_PatchData
    pTarget             As PTR
    pFailedBranch       As PTR
    pContinueBranch     As PTR
    pLastTokenID        As PTR
    pCurTokenStr        As PTR
    pExtractToken       As PTR
    lOriginalOffset     As Long
    bPatched            As Boolean
End Type

' // This structure hold data that ensures proper translation PCR->BSCR with CDecl CC
Private Type tBSCRCheckCC_PatchData
    pTarget             As PTR
    pContinue           As PTR
    bOriginalData(4)    As Byte
    bPatched            As Boolean
End Type

' // This structure hold data that ensures displaying CDecl keyword in code pane
Private Type tCodePaneDisp_PatchData
    pTarget             As PTR
    iOriginalData       As Integer
    bPatched            As Boolean
End Type

' // This structure hold data that ensures proper P-code procedure description building
Private Type tPCodeProcDesc_PatchData
    pTarget             As PTR
    pContinue           As PTR
    bOriginalData(4)    As Byte
    bPatched            As Boolean
End Type

Private m_cVba6ModInfo          As CPEModuleInfo
Private m_cThunksLib            As CThunksLibrary
Private m_tPCRIdCheckPatch      As tPCRCheckId_PatchData
Private m_tBSCRCheckCCPatch     As tBSCRCheckCC_PatchData
Private m_tCodePaneDispPatch    As tCodePaneDisp_PatchData
Private m_tPCodeProcDectPatch   As tPCodeProcDesc_PatchData
Private m_bIsInitialized        As Boolean

Private Sub Class_Initialize()

    Set m_cVba6ModInfo = New CPEModuleInfo
    
    m_cVba6ModInfo.LoadAndLock "vba6"
    
End Sub

Private Sub Class_Terminate()
    
    Patch_internal False

    m_bIsInitialized = False
    
    Set m_cThunksLib = Nothing
    
End Sub

' //
' // IImportProvider implementation
' //

Private Property Get IImportProvider_Exists( _
                     ByRef sName As String) As Boolean
    
    Select Case sName
    Case "_imp_LastTokenId", "_imp_CurrentTokenStr", "_imp_ExtractCurrentToken", "_imp_ContinueParseId", "_imp_ParseFailed", _
         "_imp_ContinueBSCRBuild", "_imp_ContinueProcDescBuild"
        IImportProvider_Exists = True
    End Select
    
End Property

Private Property Get IImportProvider_Item( _
                     ByRef sName As String) As PTR
    Const PROC_NAME = "IImportProvider_Item", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
    
    Select Case sName
    Case "_imp_LastTokenId":            IImportProvider_Item = m_tPCRIdCheckPatch.pLastTokenID
    Case "_imp_CurrentTokenStr":        IImportProvider_Item = m_tPCRIdCheckPatch.pCurTokenStr
    Case "_imp_ExtractCurrentToken":    IImportProvider_Item = m_tPCRIdCheckPatch.pExtractToken
    Case "_imp_ContinueParseId":        IImportProvider_Item = m_tPCRIdCheckPatch.pContinueBranch
    Case "_imp_ParseFailed":            IImportProvider_Item = m_tPCRIdCheckPatch.pFailedBranch
    Case "_imp_ContinueBSCRBuild":      IImportProvider_Item = m_tBSCRCheckCCPatch.pContinue
    Case "_imp_ContinueProcDescBuild":  IImportProvider_Item = m_tPCodeProcDectPatch.pContinue
    Case Else
        Err.Raise 5, PROC_NAME
    End Select
    
    Exit Property

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Property

Private Sub IPatchEntry_Initialize( _
            ByVal cScanner As CSignaturesScanner)
    Const PROC_NAME = "IPatchEntry_Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim tPCR    As tPCRCheckId_PatchData
    Dim tBSCR   As tBSCRCheckCC_PatchData
    Dim tCPDisp As tCodePaneDisp_PatchData
    Dim tPCGen  As tPCodeProcDesc_PatchData
    Dim cThunks As CThunksLibrary
    
    On Error GoTo err_handler
    
    tPCR = ExtractPCRPatchData(cScanner)
    tBSCR = ExtractBSCRPatchData(cScanner)
    tCPDisp = ExtractCodePaneDispPatchData(cScanner)
    tPCGen = ExtractPCodeDescPatchData(cScanner)
    
    Set cThunks = New CThunksLibrary
    
    ' // Load thunks library
    cThunks.Initialize LoadResData(RID_THUNKS, RT_RCDATA)
    
    m_tPCRIdCheckPatch = tPCR
    m_tBSCRCheckCCPatch = tBSCR
    m_tCodePaneDispPatch = tCPDisp
    m_tPCodeProcDectPatch = tPCGen
    
    cThunks.ResolveImport Me
    
    If cThunks.HasUnresolvedImport Then
        Err.Raise 5, , "Has unresolved import"
    End If
    
    Set m_cThunksLib = cThunks
    
    m_bIsInitialized = True
    
    Exit Sub
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description

End Sub

Private Function ExtractPCodeDescPatchData( _
                 ByVal cScanner As CSignaturesScanner) As tPCodeProcDesc_PatchData
    Const PROC_NAME = "ExtractPCodeDescPatchData", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pTest   As PTR
    
    On Error GoTo err_handler

    ' // Set specified sequence
    cScanner.StartSearch "PCode_Procedure_Desc_Patch"

    pTest = cScanner.SearchNext()
    
    Debug.Print Hex$(pTest)
    
    If pTest = NULL_PTR Then
        Err.Raise 5, , "Not found"
    End If
        
    ExtractPCodeDescPatchData.pTarget = cScanner.Target("PatchPos")
    ExtractPCodeDescPatchData.pContinue = cScanner.Target("Continue")
    
    GetMem4 ByVal ExtractPCodeDescPatchData.pTarget, ExtractPCodeDescPatchData.bOriginalData(0)
    GetMem1 ByVal ExtractPCodeDescPatchData.pTarget + 4, ExtractPCodeDescPatchData.bOriginalData(4)
    
    Exit Function
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
        
End Function

Private Function ExtractCodePaneDispPatchData( _
                 ByVal cScanner As CSignaturesScanner) As tCodePaneDisp_PatchData
    Const PROC_NAME = "ExtractCodePaneDispPatchData", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pTest   As PTR
    
    On Error GoTo err_handler

    ' // Set specified sequence
    cScanner.StartSearch "CodePane_DisplayPatch"

    pTest = cScanner.SearchNext()
    
    Debug.Print Hex$(pTest)
    
    If pTest = NULL_PTR Then
        Err.Raise 5, , "Not found"
    End If
        
    ExtractCodePaneDispPatchData.pTarget = cScanner.Target("PatchPos")
    
    GetMem2 ByVal ExtractCodePaneDispPatchData.pTarget, ExtractCodePaneDispPatchData.iOriginalData
    
    Exit Function
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
        
End Function

Private Function ExtractBSCRPatchData( _
                 ByVal cScanner As CSignaturesScanner) As tBSCRCheckCC_PatchData
    Const PROC_NAME = "ExtractBSCRPatchData", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pTest   As PTR
    
    On Error GoTo err_handler

    ' // Set specified sequence
    cScanner.StartSearch "BSCR_Check_CC"

    pTest = cScanner.SearchNext()
    
    Debug.Print Hex$(pTest)
    
    If pTest = NULL_PTR Then
        Err.Raise 5, , "Not found"
    End If
        
    ExtractBSCRPatchData.pTarget = cScanner.Target("PatchPos")
    ExtractBSCRPatchData.pContinue = cScanner.Target("Continue")
    
    GetMem4 ByVal ExtractBSCRPatchData.pTarget, ExtractBSCRPatchData.bOriginalData(0)
    GetMem1 ByVal ExtractBSCRPatchData.pTarget + 4, ExtractBSCRPatchData.bOriginalData(4)
    
    Exit Function
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
        
End Function

Private Function ExtractPCRPatchData( _
                 ByVal cScanner As CSignaturesScanner) As tPCRCheckId_PatchData
    Const PROC_NAME = "ExtractPCRPatchData", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pTest           As PTR
    Dim pLastID         As PTR
    Dim pCurTokenStr    As PTR
    Dim pFailedBranch   As PTR
    Dim pContinueBranch As PTR
    Dim pExtractToken   As PTR
    Dim tSecText        As IMAGE_SECTION_HEADER
    Dim tSecData        As IMAGE_SECTION_HEADER
    
    On Error GoTo err_handler
    
    ' // Check ranges
    tSecText = m_cVba6ModInfo.SectionByName(".text")
    tSecData = m_cVba6ModInfo.SectionByName(".data")
    
    ' // Set specified sequence
    cScanner.StartSearch "PCR_Check_Id"

    Do
        
        pTest = cScanner.SearchNext()
        
        Debug.Print Hex$(pTest)
        
        If pTest = NULL_PTR Then
            Err.Raise 5, , "Not found"
        End If
        
        pLastID = cScanner.Target("LastTokenID")
        pCurTokenStr = cScanner.Target("CurrenctTokenStr")
        pFailedBranch = cScanner.Target("ParseFailed")
        pContinueBranch = cScanner.Target("ParseIdContinue")
        pExtractToken = cScanner.Target("ExtractCurrentToken")
        
        If pLastID >= tSecData.VirtualAddress + m_cVba6ModInfo.BaseAddress And _
            pLastID < tSecData.VirtualAddress + tSecData.VirtualSize + m_cVba6ModInfo.BaseAddress And _
            pCurTokenStr >= tSecData.VirtualAddress + m_cVba6ModInfo.BaseAddress And _
            pCurTokenStr < tSecData.VirtualAddress + tSecData.VirtualSize + m_cVba6ModInfo.BaseAddress And _
            pFailedBranch >= tSecText.VirtualAddress + m_cVba6ModInfo.BaseAddress And _
            pFailedBranch < tSecText.VirtualAddress + tSecText.VirtualSize + m_cVba6ModInfo.BaseAddress And _
            pContinueBranch >= tSecText.VirtualAddress + m_cVba6ModInfo.BaseAddress And _
            pContinueBranch < tSecText.VirtualAddress + tSecText.VirtualSize + m_cVba6ModInfo.BaseAddress And _
            pExtractToken >= tSecText.VirtualAddress + m_cVba6ModInfo.BaseAddress And _
            pExtractToken < tSecText.VirtualAddress + tSecText.VirtualSize + m_cVba6ModInfo.BaseAddress Then
        
            ExtractPCRPatchData.pTarget = cScanner.Target("PatchPos")
            ExtractPCRPatchData.pFailedBranch = pFailedBranch
            ExtractPCRPatchData.pContinueBranch = pContinueBranch
            ExtractPCRPatchData.pLastTokenID = pLastID
            ExtractPCRPatchData.pCurTokenStr = pCurTokenStr
            ExtractPCRPatchData.pExtractToken = pExtractToken
            
            GetMem4 ByVal ExtractPCRPatchData.pTarget, ExtractPCRPatchData.lOriginalOffset

            Exit Function
            
        End If
        
    Loop While True

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
    
    Exit Sub
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
         
End Sub

Private Sub IPatchEntry_Unpatch()
    Const PROC_NAME = "IPatchEntry_Unpatch", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
    
    If Not m_bIsInitialized Then
        Err.Raise 5, , "Not initialized"
    End If
    
    Patch_internal False
    
    Exit Sub
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
         
End Sub

Private Sub Patch_internal( _
            ByVal bPatch As Boolean)
    Const PROC_NAME = "Patch_internal", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bJmpBuffer(5)   As Byte
    
    bJmpBuffer(0) = &HE9
    
    If m_tPCRIdCheckPatch.bPatched <> bPatch Then
    
        If bPatch Then
            WriteMem m_tPCRIdCheckPatch.pTarget, _
                     VarPtr(m_cThunksLib.Thunk("PCR_ParseThunk") - (m_tPCRIdCheckPatch.pTarget + 4)), 4
        Else
            WriteMem m_tPCRIdCheckPatch.pTarget, VarPtr(m_tPCRIdCheckPatch.lOriginalOffset), 4
        End If
        
        m_tPCRIdCheckPatch.bPatched = bPatch
        
    End If
    
    If m_tBSCRCheckCCPatch.bPatched <> bPatch Then
        
        If bPatch Then
        
            PutMem4 bJmpBuffer(1), m_cThunksLib.Thunk("BSCR_CreateThunk") - (m_tBSCRCheckCCPatch.pTarget + 5)
            WriteMem m_tBSCRCheckCCPatch.pTarget, VarPtr(bJmpBuffer(0)), 5

        Else
            WriteMem m_tBSCRCheckCCPatch.pTarget, VarPtr(m_tBSCRCheckCCPatch.bOriginalData(0)), 5
        End If
        
        m_tBSCRCheckCCPatch.bPatched = bPatch
    
    End If
    
    If m_tCodePaneDispPatch.bPatched <> bPatch Then
        
        If bPatch Then
            WriteMem m_tCodePaneDispPatch.pTarget, VarPtr(&H9090), 2
        Else
            WriteMem m_tCodePaneDispPatch.pTarget, VarPtr(m_tCodePaneDispPatch.iOriginalData), 2
        End If
        
        m_tCodePaneDispPatch.bPatched = bPatch
    
    End If
    
    If m_tPCodeProcDectPatch.bPatched <> bPatch Then
        
        If bPatch Then
        
            PutMem4 bJmpBuffer(1), m_cThunksLib.Thunk("ProcDesc_Thunk") - (m_tPCodeProcDectPatch.pTarget + 5)
            WriteMem m_tPCodeProcDectPatch.pTarget, VarPtr(bJmpBuffer(0)), 5

        Else
            WriteMem m_tPCodeProcDectPatch.pTarget, VarPtr(m_tPCodeProcDectPatch.bOriginalData(0)), 5
        End If
        
        m_tPCodeProcDectPatch.bPatched = bPatch
    
    End If
    
End Sub

Private Sub WriteMem( _
            ByVal pDest As PTR, _
            ByVal pSrc As PTR, _
            ByVal szSize As SIZE_T)
    Const PROC_NAME = "Patch_internal", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
                
    Dim lOldProtect As Long
    
    If VirtualProtect(pDest, szSize, PAGE_EXECUTE_READWRITE, lOldProtect) = 0 Then
        Err.Raise 7, PROC_NAME
    End If
    
    Select Case szSize
    Case 1:     GetMem1 ByVal pSrc, ByVal pDest
    Case 2:     GetMem2 ByVal pSrc, ByVal pDest
    Case 4:     GetMem4 ByVal pSrc, ByVal pDest
    Case 8:     GetMem8 ByVal pSrc, ByVal pDest
    Case Else:  memcpy ByVal pDest, ByVal pSrc, szSize
    End Select
    
    VirtualProtect pDest, szSize, lOldProtect, lOldProtect
                
End Sub


