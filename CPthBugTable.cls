VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPthBugTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CPthBugTable.cls - serves patch to fix crashes when cdecl function declared in a typelibrary
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME   As String = "CPthBugTable"

Implements IPatchEntry

Private m_cVba6ModInfo      As CPEModuleInfo
Private m_pTable            As PTR
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
    
    Dim pTable      As PTR
    Dim tSecRange   As IMAGE_SECTION_HEADER
    Dim bBugValue   As Byte
    
    On Error GoTo err_handler
    
    cScanner.StartSearch "Bug_Signature"
    
    tSecRange = m_cVba6ModInfo.SectionByName(".text")
    
    Do
        
        ' // Search for code sequence where table is referenced
        If cScanner.SearchNext() = NULL_PTR Then
            Err.Raise 5, vbNullString, "Not found"
        End If
        
        ' // Get pointer to address to table
        pTable = cScanner.Target("BugTablePtr")

        ' // Check if address is within .text section
        If pTable >= tSecRange.VirtualAddress + m_cVba6ModInfo.BaseAddress Or _
            pTable + &H1D < tSecRange.VirtualAddress + tSecRange.VirtualSize + m_cVba6ModInfo.BaseAddress Then
            
            ' // Check value
            GetMem1 ByVal pTable + &H1D, bBugValue
            
            If bBugValue = &H19 Then
                
                m_bIsInitialized = True
                m_pTable = pTable
                Exit Sub
                
            End If
            
        End If
        
    Loop While True
    
    Exit Sub

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
    Dim bNewValue   As Byte
    
    If VirtualProtect(m_pTable, &H1E, PAGE_EXECUTE_READWRITE, lOldProtect) = 0 Then
        Err.Raise 7, PROC_NAME
    End If

    If bPatch Then
        bNewValue = 9
    Else
        bNewValue = &H19
    End If
    
    PutMem1 ByVal m_pTable + &H1D, bNewValue
    
    VirtualProtect m_pTable, &H1E, lOldProtect, lOldProtect
        
End Sub




