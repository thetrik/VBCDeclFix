Attribute VB_Name = "modMain"
' //
' // modMain.bas - global functions for VBCdeclFix project
' // by The trick 2021
' //

Option Explicit

Public Enum eResourcesIDs
    
    RID_OPCODES_HANDLERS = 101
    RID_SIGNATURES = 102
    RID_THUNKS = 103
    
End Enum

Private Const MODULE_NAME   As String = "modMain"

' //
' // Get number of elements in array by pointer
' //
Public Function SafeArrayElementsCount( _
                ByVal ppSA As PTR) As Long
    Dim tSA     As SAFEARRAY1D
    Dim pSA     As PTR
    Dim pBound  As PTR
    Dim tBound  As SAFEARRAYBOUND
    
    If ppSA = 0 Then Exit Function
    
    GetMemPtr ByVal ppSA, pSA
    
    If pSA = 0 Then Exit Function
    
    memcpy tSA, ByVal pSA, Len(tSA) - Len(tBound)
    
    pBound = pSA + Len(tSA) - Len(tBound)
    
    SafeArrayElementsCount = 1
    
    Do While tSA.cDims > 0
    
        memcpy tBound, ByVal pBound, Len(tBound)
        
        SafeArrayElementsCount = SafeArrayElementsCount * tBound.cElements
        pBound = pBound + Len(tBound)
        tSA.cDims = tSA.cDims - 1
        
    Loop
                    
End Function

