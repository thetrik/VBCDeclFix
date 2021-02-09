VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSignatureSearcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CSignatureSearcher.cls - search a pattern in memory using mask
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME   As String = "CSignatureSearcher"

Private m_bPattern()    As Byte
Private m_bMask()       As Byte
Private m_lPatternSize  As Long

' // Initialize
Public Sub Initialize( _
           ByVal pPattern As PTR, _
           ByVal lPatternSize As Long, _
           ByVal pMask As PTR, _
           ByVal lMaskSize As Long)
    Const PROC_NAME = "Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
    
    If lPatternSize <= 0 Or pPattern = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME, "Invalid pattern"
    End If
    
    If lMaskSize < 0 Or lMaskSize > lPatternSize Or (pMask = NULL_PTR And lMaskSize <> 0) Then
        Err.Raise 5, FULL_PROC_NAME, "Invalid mask"
    End If
    
    ReDim m_bPattern(lPatternSize - 1)
    ReDim m_bMask(lPatternSize - 1)
    
    memcpy m_bPattern(0), ByVal pPattern, lPatternSize
    
    If lMaskSize Then
        memcpy m_bMask(0), ByVal pMask, lMaskSize
    End If
    
    If lMaskSize < lPatternSize Then
        memset m_bMask(lMaskSize), lPatternSize - lMaskSize, &HFF
    End If
    
    m_lPatternSize = lPatternSize
    
    Exit Sub
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
                
End Sub

' // Search current pattern
Public Function Search( _
                ByVal pStartAddr As PTR, _
                ByVal lSize As Long) As PTR
    Const PROC_NAME = "Search", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
        
    Dim tSADesc As SAFEARRAY1D
    Dim bData() As Byte
    Dim lN      As Long
    Dim lZ      As Long
    
    With tSADesc
        .cbElements = 1
        .cDims = 1
        .fFeatures = FADF_AUTO
        .pvData = pStartAddr
        .rgsabound(0).cElements = lSize
    End With
    
    vbaAryMove ByVal ArrPtr(bData), VarPtr(tSADesc)

    Do While lN < lSize - m_lPatternSize

        If (bData(lN) And m_bMask(lZ)) = (m_bPattern(lZ) And m_bMask(lZ)) Then
        
            lZ = lZ + 1
            
            If lZ >= m_lPatternSize Then
                Search = lN - (m_lPatternSize - 1) + pStartAddr
                Exit Do
            End If
            
        Else
        
            If lZ Then
                lN = lN - lZ + 1
                lZ = 0
            End If
            
        End If
        
        lN = lN + 1
        
    Loop
    
    Exit Function
    
err_handler:
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
      
End Function

Public Property Get PatternSize() As Long
    PatternSize = m_lPatternSize
End Property

Public Property Get Pattern() As Byte()
    Pattern = m_bPattern
End Property

Public Property Get Mask() As Byte()
    Mask = m_bMask
End Property


