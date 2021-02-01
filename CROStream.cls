VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CROStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CROStream.cls - read only stream
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME       As String = "CROStream"

Private m_bData()   As Byte
Private m_lOffset   As Long
Private m_lDataSize As Long

Public Sub InitializeFromByteArray( _
           ByRef bData() As Byte)
               
    If SafeArrayElementsCount(ArrPtr(bData)) Then
        Initialize VarPtr(bData(0)), UBound(bData) + 1
    Else
        Initialize NULL_PTR, 0
    End If
               
End Sub

Public Sub Initialize( _
           ByVal pData As PTR, _
           ByVal lSize As Long)
    Const PROC_NAME = "Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If (pData = 0 And lSize <> 0) Or lSize < 0 Then
        Err.Raise 5, FULL_PROC_NAME, "Invalid data"
    End If
    
    If lSize > 0 Then
        
        ReDim m_bData(lSize - 1)
        
        memcpy m_bData(0), ByVal pData, lSize
        
    Else
        Erase m_bData
    End If
    
    m_lOffset = 0
    m_lDataSize = lSize
    
End Sub

Public Function ExtractPart( _
                ByVal lOffset As Long, _
                ByVal lSize As Long) As CROStream
    Const PROC_NAME = "ExtractPart", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
          
    If lOffset + lSize > m_lDataSize Or lOffset < 0 Or lSize <= 0 Then
        Err.Raise 9, FULL_PROC_NAME
    End If
               
    Set ExtractPart = New CROStream
    
    ExtractPart.Initialize VarPtr(m_bData(lOffset)), lSize
               
End Function

Public Property Get EOF() As Boolean
    EOF = m_lOffset = m_lDataSize
End Property

Public Sub Align( _
           ByVal lValue As Long)
    Const PROC_NAME = "Align", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lNewPos As Long
    
    lNewPos = ((m_lOffset + (lValue - 1)) \ lValue) * lValue
    
    If lNewPos > m_lDataSize Then
        Err.Raise 9, FULL_PROC_NAME
    End If
            
    m_lOffset = lNewPos
    
End Sub

Public Function ReadByte() As Byte
    Const PROC_NAME = "ReadByte", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_lOffset + Len(ReadByte) > m_lDataSize Then
        Err.Raise 9, FULL_PROC_NAME
    End If
    
    ReadByte = m_bData(m_lOffset)
    
    m_lOffset = m_lOffset + Len(ReadByte)
    
End Function

Public Function ReadInteger() As Integer
    Const PROC_NAME = "ReadInteger", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_lOffset + Len(ReadInteger) > m_lDataSize Then
        Err.Raise 9, FULL_PROC_NAME
    End If
    
    GetMem1 m_bData(m_lOffset), ReadInteger
    
    m_lOffset = m_lOffset + Len(ReadInteger)
    
End Function

Public Function ReadLong() As Long
    Const PROC_NAME = "ReadLong", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_lOffset + Len(ReadLong) > m_lDataSize Then
        Err.Raise 9, FULL_PROC_NAME
    End If
    
    GetMem4 m_bData(m_lOffset), ReadLong
    
    m_lOffset = m_lOffset + Len(ReadLong)
    
End Function

Public Function ReadMBZStringToUTF16() As String
    Const PROC_NAME = "ReadMBZString", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lIndex  As Long
    Dim lSize   As Long
    
    For lIndex = m_lOffset To m_lDataSize - 1
        If m_bData(lIndex) = 0 Then
            Exit For
        End If
    Next
    
    If lIndex = m_lDataSize Then
        Err.Raise 9, FULL_PROC_NAME
    End If
    
    lSize = MultiByteToWideChar(CP_ACP, 0, m_bData(m_lOffset), lIndex - m_lOffset, ByVal NULL_PTR, 0)
    
    If lSize = 0 Then
        Err.Raise 7, FULL_PROC_NAME
    End If
    
    ReadMBZStringToUTF16 = Space$(lSize)
    
    If MultiByteToWideChar(CP_ACP, 0, m_bData(m_lOffset), lIndex - m_lOffset, ByVal StrPtr(ReadMBZStringToUTF16), lSize) = 0 Then
        Err.Raise 7, FULL_PROC_NAME
    End If
    
    m_lOffset = lIndex + 1
    
End Function

Public Sub ReadData( _
           ByVal pTo As PTR, _
           ByVal lSize As Long)
    Const PROC_NAME = "ReadData", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_lOffset + lSize > m_lDataSize Then
        Err.Raise 9, FULL_PROC_NAME
    End If
     
    memcpy ByVal pTo, m_bData(m_lOffset), lSize
    
    m_lOffset = m_lOffset + lSize
     
End Sub

Public Property Get Size() As Long
    Size = m_lDataSize
End Property

' // Default (Procedure attributes)
Public Property Get Position() As Long
Attribute Position.VB_UserMemId = 0
    Position = m_lOffset
End Property

Public Property Let Position( _
                    ByVal lValue As Long)
    Const PROC_NAME = "Position_put", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If lValue < 0 Or lValue > m_lDataSize Then
        Err.Raise 5, FULL_PROC_NAME, "Invalid pos"
    End If
    
    m_lOffset = lValue
    
End Property

