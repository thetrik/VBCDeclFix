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

Private Const MODULE_NAME   As String = "CRuntimePatcher"

Private m_cPatches          As Collection
Private m_bIsInitialized    As Boolean
Private m_bCDeclEnabled     As Boolean

' // Initialize module
Public Sub Initialize()
    Const PROC_NAME = "Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim cScanner    As CSignaturesScanner
    Dim lIndex      As Long
    Dim cPatch      As IPatchEntry
    Dim cPatches    As Collection
    
    On Error GoTo err_handler
    
    If m_bIsInitialized Then
        Exit Sub
    End If
    
    Set cScanner = New CSignaturesScanner
    
    ' // Initialize signatures from resource
    cScanner.InitializeFromByteArray LoadResData(RID_SIGNATURES, RT_RCDATA)
    
    Set cPatches = New Collection
    
    For lIndex = 0 To 3
        
        Select Case lIndex
        Case 0: Set cPatch = New CPthBugTable
        Case 1: Set cPatch = New CPthOpHandlers
        Case 2: Set cPatch = New CPthDeclareFIx
        Case 3: Set cPatch = New CPthUserCDeclFn
        End Select
        
        cPatch.Initialize cScanner
        
        cPatches.Add cPatch
        
    Next
    
    Set m_cPatches = cPatches
    
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

    Dim cPatch  As IPatchEntry
    
    On Error GoTo err_handler

    For Each cPatch In m_cPatches
        cPatch.Unpatch
    Next

    m_bCDeclEnabled = False
    
    Exit Sub

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Sub

' // Enable CDECL
Private Sub EnableCdecl()
    Const PROC_NAME = "EnableCdecl", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    Dim cPatch  As IPatchEntry
    
    On Error GoTo err_handler

    For Each cPatch In m_cPatches
        cPatch.Patch
    Next
    
    m_bCDeclEnabled = True
    
    Exit Sub

err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Sub

Private Sub Class_Terminate()
    
    If m_bCDeclEnabled Then
        DisableCdecl
    End If
    
    Set m_cPatches = Nothing
    
    m_bIsInitialized = False
    
End Sub




