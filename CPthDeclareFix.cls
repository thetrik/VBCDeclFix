VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPthDeclareFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CPthDeclareFix.cls - serves patch to enable CDecl keyword in Declare statement
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME   As String = "CPthDeclareFix"

Implements IPatchEntry

Private m_cVba6ModInfo      As CPEModuleInfo
Private m_pPatchTarget      As PTR
Private m_pJmpTo            As PTR
Private m_lOriginalOffset   As Long
Private m_bPatched          As Boolean
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
    
End Sub

Private Sub IPatchEntry_Initialize( _
            ByVal cScanner As CSignaturesScanner)
    Const PROC_NAME = "IPatchEntry_Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
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
            Err.Raise 5, , "Not found"
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
                
                m_pPatchTarget = pJmpTo
                m_pJmpTo = pTarget
                
                GetMem4 ByVal pJmpTo, m_lOriginalOffset
                
                m_bIsInitialized = True
                
                Exit Sub
                
            End If
            
        End If
        
    Loop While True

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
         
End Sub

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
    
    If Not m_bIsInitialized Then
        Err.Raise 5, , "Not initialized"
    End If
    
    Patch_internal False
    m_bPatched = False
    
    Exit Sub
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
         
End Sub

Private Sub Patch_internal( _
            ByVal bPatch As Boolean)
    Const PROC_NAME = "Patch_internal", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lOldProtect As Long
    Dim lOffset     As Long
    
    If VirtualProtect(m_pPatchTarget, &H4, PAGE_EXECUTE_READWRITE, lOldProtect) = 0 Then
        Err.Raise 7, PROC_NAME
    End If

    If bPatch Then
        lOffset = m_pJmpTo - (m_pPatchTarget + 4)
    Else
        lOffset = m_lOriginalOffset
    End If
    
    PutMem4 ByVal m_pPatchTarget, lOffset
    
    VirtualProtect m_pPatchTarget, &H4, lOldProtect, lOldProtect
    
End Sub


