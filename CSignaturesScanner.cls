VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSignaturesScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CSignaturesScanner.cls - signature scanner in memory using desc file
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME   As String = "CSignaturesScanner"

' // File item tag
' // Each item has the following structure
' // 0 - BYTE - Tag
' // 1 - DWORD - Variable data size
' // 5 - VARIABLE - Data
Private Enum eSignFileTags
    TAG_MODULE = 1
    TAG_SIGNATURE = 2
    TAG_RANGE_LIST = 3
    TAG_MODULE_RANGE = 4
    TAG_PATTERN = 5
    TAG_MASK = 6
    TAG_TARGET = 7
End Enum

Private Enum eTargetType
    TARGET_ABSOLUTE = 0      ' Target is offset itself
    TARGET_PTR = 1           ' Target is pointer to pointer
    TARGET_RELATIVE = 2      ' Target is value relative target address (call/jmp DWORD)
End Enum

Private Type tSignatureTarget
    eType           As eTargetType
    sName           As String
    lOffset         As Long
End Type

Private Type tModuleRange
    sSectionName    As String
    lStartOffset    As Long
    lEndOffset      As Long
End Type

Private Type tSignature
    sName           As String
    bPattern()      As Byte
    bMask()         As Byte
    tRanges()       As tModuleRange
    tTargets()      As tSignatureTarget
End Type

Private Type tModule
    sName           As String
    tSignatures()   As tSignature
End Type

Private Type tSignaturePath
    lModuleIndex    As Long
    lSignatureIndex As Long
End Type

Private Type tSearchRange
    pStart          As PTR
    szSize          As SIZE_T
End Type

Private m_tModules()        As tModule
Private m_lSignaturesCount  As Long
Private m_cSearcher         As CSignatureSearcher
Private m_tCurSignaturePath As tSignaturePath
Private m_cCurModule        As CPEModuleInfo
Private m_tCurRange         As tSearchRange
Private m_lCurRangeIndex    As Long
Private m_pLastFound        As PTR

Public Sub InitializeFromByteArray( _
           ByRef bData() As Byte)
    Const PROC_NAME = "InitializeFromByteArray", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If SafeArrayElementsCount(ArrPtr(bData)) Then
        Initialize VarPtr(bData(0)), UBound(bData) + 1
    Else
        Err.Raise 5, FULL_PROC_NAME
    End If
               
End Sub

' // Load definitions from file (in memory)
Public Sub Initialize( _
           ByVal pData As PTR, _
           ByVal lDataSize As Long)
    Const PROC_NAME = "Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
    
    Dim cStream     As CROStream
    Dim eTag        As eSignFileTags
    Dim tResult()   As tModule
    Dim lResCount   As Long
    Dim lSize       As Long
    Dim lIndex      As Long
    Dim lSignsCount As Long
    Dim lIndex2     As Long
    
    Set cStream = New CROStream
    
    cStream.Initialize pData, lDataSize
    
    Do Until cStream.EOF
        
        eTag = cStream.ReadByte()
        
        Select Case eTag
        Case TAG_MODULE
            
            If lResCount Then
                If lResCount > UBound(tResult) Then
                    ReDim Preserve tResult(lResCount + 10)
                End If
            Else
                ReDim tResult(9)
            End If
            
            tResult(lResCount) = ExtractModuleDefinition(cStream)
            
            lResCount = lResCount + 1
            
        Case Else
            ' // Skip
            
            lSize = cStream.ReadLong()
            cStream.Position = cStream.Position + lSize
            
        End Select
        
        cStream.Align 4 ' // Alignment required
        
    Loop
    
    If lResCount Then
        ReDim Preserve tResult(lResCount - 1)
    Else
        Erase tResult
    End If
    
    For lIndex = 0 To lResCount - 1
        
        lSize = SafeArrayElementsCount(ArrPtr(tResult(lIndex).tSignatures))
        
        For lIndex2 = 0 To lSize - 1
            
            If SafeArrayElementsCount(ArrPtr(tResult(lIndex).tSignatures(lIndex2).bPattern)) <= 0 Or _
                SafeArrayElementsCount(ArrPtr(tResult(lIndex).tSignatures(lIndex2).tRanges)) <= 0 Or _
                SafeArrayElementsCount(ArrPtr(tResult(lIndex).tSignatures(lIndex2).tTargets)) <= 0 Or _
                Len(tResult(lIndex).tSignatures(lIndex2).sName) = 0 Then
                
                Err.Raise 5, , "Invalid format"
                
            End If
            
        Next
        
        lSignsCount = lSignsCount + lSize
        
    Next
    
    m_tModules = tResult
    m_lSignaturesCount = lSignsCount
    
    ReleaseSearch
    
    Exit Sub
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
           
End Sub

Public Property Get SignaturesCount() As Long
    SignaturesCount = m_lSignaturesCount
End Property

Public Property Get SignatureName( _
                    ByVal lIndex As Long) As String
    Const PROC_NAME = "SignatureName_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
        
    Dim lModIndex   As Long
    Dim lStartIndex As Long
    Dim lCount      As Long
    
    If lIndex < 0 Or lIndex >= m_lSignaturesCount Then
        Err.Raise 9, FULL_PROC_NAME
    End If
    
    For lModIndex = 0 To UBound(m_tModules)
        
        lCount = SafeArrayElementsCount(ArrPtr(m_tModules(lModIndex).tSignatures))
        
        If lStartIndex + lCount > lIndex Then
            SignatureName = m_tModules(lModIndex).tSignatures(lIndex - lStartIndex).sName
            Exit Property
        End If
        
        lStartIndex = lStartIndex + lCount
        
    Next
 
End Property

Public Property Get SignatureExists( _
                    ByRef sName As String) As Boolean
    Const PROC_NAME = "SignatureExists_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
        
    Dim lModIndex   As Long
    Dim lSignature  As Long

    For lModIndex = 0 To UBound(m_tModules)
        For lSignature = 0 To UBound(m_tModules(lModIndex).tSignatures)
            If StrComp(sName, m_tModules(lModIndex).tSignatures(lSignature).sName, vbBinaryCompare) = 0 Then
                
                SignatureExists = True
                Exit Property
                
            End If
        Next
    Next
    
End Property

Public Property Get Targets() As Collection
    Const PROC_NAME = "Targets_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim cRet        As Collection
    Dim lModIndex   As Long
    Dim lSignature  As Long
    Dim lTarget     As Long
    
    Set cRet = New Collection
    
    For lModIndex = 0 To UBound(m_tModules)
        For lSignature = 0 To UBound(m_tModules(lModIndex).tSignatures)
            For lTarget = 0 To UBound(m_tModules(lModIndex).tSignatures(lSignature).tTargets)
                cRet.Add m_tModules(lModIndex).tSignatures(lSignature).tTargets(lTarget).sName
            Next
        Next
    Next
    
    Set Targets = cRet
    
End Property

' // Get list of current targets
Public Property Get CurrentTargets() As Collection
    Const PROC_NAME = "CurrentTargets_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim cRet    As Collection
    Dim lIndex  As Long
    
    If m_pLastFound = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    Set cRet = New Collection
    
    With m_tModules(m_tCurSignaturePath.lModuleIndex).tSignatures(m_tCurSignaturePath.lSignatureIndex)
        
        For lIndex = 0 To UBound(.tTargets)
            cRet.Add .tTargets(lIndex).sName
        Next
        
    End With
    
    Set CurrentTargets = cRet
    
End Property

' // Get target address by its name
Public Property Get Target( _
                    ByRef vNameIndex As Variant) As PTR
    Const PROC_NAME = "Target_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lIndex          As Long
    Dim sName           As String
    Dim pTargetOffset   As PTR
    Dim lOffset         As Long
    
    If m_pLastFound = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    With m_tModules(m_tCurSignaturePath.lModuleIndex).tSignatures(m_tCurSignaturePath.lSignatureIndex)
        
        Select Case VarType(vNameIndex)
        Case vbLong, vbInteger, vbByte
        
            lIndex = CLng(vNameIndex)
            
            If lIndex < 0 Or lIndex > UBound(.tTargets) Then
                Err.Raise 9, FULL_PROC_NAME
            End If
            
            pTargetOffset = m_pLastFound + .tTargets(lIndex).lOffset

        Case Else
            
            sName = CStr(vNameIndex)
            
            For lIndex = 0 To UBound(.tTargets)
                
                If StrComp(.tTargets(lIndex).sName, sName, vbBinaryCompare) = 0 Then
                    
                    pTargetOffset = m_pLastFound + .tTargets(lIndex).lOffset
                    Exit For
                    
                End If
                
            Next
        
        End Select
        
        If pTargetOffset = NULL_PTR Then
            Err.Raise 5, FULL_PROC_NAME
        End If
        
        Select Case .tTargets(lIndex).eType
        Case TARGET_ABSOLUTE
            
            Target = pTargetOffset
        
        Case TARGET_PTR
        
            GetMemPtr ByVal pTargetOffset, Target
        
        Case TARGET_RELATIVE
        
            GetMem4 ByVal pTargetOffset, lOffset
            Target = pTargetOffset + 4 + lOffset
        
        End Select
        
    End With
    
End Property

Public Sub StartSearch( _
           ByRef sSignatureName As String)
    Const PROC_NAME = "StartSearch", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    Dim tSignPath   As tSignaturePath
    Dim cSearcher   As CSignatureSearcher
    Dim cModule     As CPEModuleInfo
    Dim tRange      As tSearchRange
    
    On Error GoTo err_handler

    tSignPath = SignatureModuleIndexByName(sSignatureName)
    
    With m_tModules(tSignPath.lModuleIndex).tSignatures(tSignPath.lSignatureIndex)

        Set cModule = New CPEModuleInfo
        
        cModule.LoadAndLock m_tModules(tSignPath.lModuleIndex).sName
        
        tRange = SearchRangeFromModuleRange(cModule, .tRanges(0))
        
        Set cSearcher = New CSignatureSearcher
    
        If SafeArrayElementsCount(ArrPtr(.bMask)) > 0 Then
            cSearcher.Initialize VarPtr(.bPattern(0)), UBound(.bPattern) + 1, VarPtr(.bMask(0)), UBound(.bMask) + 1
        Else
            cSearcher.Initialize VarPtr(.bPattern(0)), UBound(.bPattern) + 1, NULL_PTR, 0
        End If
        
        ReleaseSearch
        
        m_tCurRange = tRange
        m_tCurSignaturePath = tSignPath
        
        Set m_cCurModule = cModule
        Set m_cSearcher = cSearcher
        
    End With
    
    Exit Sub
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
      
End Sub

Public Function Match( _
                ByVal pData As PTR, _
                ByVal lSize As Long) As Boolean
    Const PROC_NAME = "Match", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    On Error GoTo err_handler
    
    If m_cSearcher Is Nothing Then
        Err.Raise 5
    End If

    Match = m_cSearcher.Search(pData, lSize) = pData

    Exit Function
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
            
End Function

Public Function SearchNext() As PTR
    Const PROC_NAME = "SearchNext", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim pNewPos As PTR
    
    On Error GoTo err_handler
    
    If m_cSearcher Is Nothing Then
        Err.Raise 5
    End If
    
    Do
    
        pNewPos = m_cSearcher.Search(m_tCurRange.pStart, m_tCurRange.szSize)
        
        If pNewPos = NULL_PTR Then
            
            With m_tModules(m_tCurSignaturePath.lModuleIndex).tSignatures(m_tCurSignaturePath.lSignatureIndex)
            
                ' // Check next region
                If m_lCurRangeIndex < UBound(.tRanges) Then
                    
                    m_lCurRangeIndex = m_lCurRangeIndex + 1
                    m_tCurRange = SearchRangeFromModuleRange(m_cCurModule, .tRanges(m_lCurRangeIndex))
                    
                Else
                    Exit Do
                End If
                  
            End With
            
        Else
            
            m_tCurRange.szSize = m_tCurRange.szSize - (pNewPos - m_tCurRange.pStart) - 1
            m_tCurRange.pStart = pNewPos + 1

        End If
        
    Loop Until pNewPos
    
    m_pLastFound = pNewPos
    SearchNext = pNewPos
    
    Exit Function
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
      
End Function

Private Sub ReleaseSearch()
    
    m_lCurRangeIndex = 0
    m_pLastFound = NULL_PTR
    m_tCurRange.pStart = 0
    m_tCurRange.szSize = 0
    
    Set m_cSearcher = Nothing
    Set m_cCurModule = Nothing
    
End Sub

Private Property Get SearchRangeFromModuleRange( _
                     ByVal cModule As CPEModuleInfo, _
                     ByRef tModRange As tModuleRange) As tSearchRange
    Const PROC_NAME = "SearchRangeFromModuleRange_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
    
    With cModule.SectionByName(tModRange.sSectionName)
        
        SearchRangeFromModuleRange.pStart = .VirtualAddress + cModule.BaseAddress + tModRange.lStartOffset
        SearchRangeFromModuleRange.szSize = .VirtualSize - tModRange.lEndOffset
    
    End With
    
    Exit Property
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
        
End Property

Private Property Get SignatureModuleIndexByName( _
                     ByRef sSignatureName As String) As tSignaturePath
    Const PROC_NAME = "SignatureModuleIndexByName_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lModIndex   As Long
    Dim lSigIndex   As Long
    
    For lModIndex = 0 To SafeArrayElementsCount(ArrPtr(m_tModules)) - 1
        For lSigIndex = 0 To SafeArrayElementsCount(ArrPtr(m_tModules(lModIndex).tSignatures)) - 1
            If StrComp(sSignatureName, m_tModules(lModIndex).tSignatures(lSigIndex).sName, vbBinaryCompare) = 0 Then
                
                SignatureModuleIndexByName.lModuleIndex = lModIndex
                SignatureModuleIndexByName.lSignatureIndex = lSigIndex
                
                Exit Property
                
            End If
        Next
    Next
    
    Err.Raise 5, FULL_PROC_NAME
    
End Property

Private Function ExtractModuleDefinition( _
                 ByVal cStm As CROStream) As tModule
    Const PROC_NAME = "ExtractModuleDefinition", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lDataSize   As Long
    Dim cModStream  As CROStream
    Dim eTag        As eSignFileTags
    Dim lSigIndex   As Long
    Dim tRet        As tModule
    
    On Error GoTo err_handler
        
    lDataSize = cStm.ReadLong()
    
    Set cModStream = cStm.ExtractPart(cStm.Position, lDataSize)
    
    cStm.Position = cStm.Position + lDataSize
    
    tRet.sName = cModStream.ReadMBZStringToUTF16()
    
    Do Until cModStream.EOF
        
        eTag = cModStream.ReadByte()
        
        Select Case eTag
        Case TAG_SIGNATURE
            
            If lSigIndex Then
                If lSigIndex > UBound(tRet.tSignatures) Then
                    ReDim tRet.tSignatures(lSigIndex + 10)
                End If
            Else
                ReDim tRet.tSignatures(9)
            End If
            
            tRet.tSignatures(lSigIndex) = ExtractSignatureDefinition(cModStream)
            
            lSigIndex = lSigIndex + 1
            
        Case Else
            ' // Skip
            
            lDataSize = cModStream.ReadLong()
            cModStream.Position = cModStream.Position + lDataSize
            
        End Select
        
    Loop
    
    If lSigIndex Then
        ReDim Preserve tRet.tSignatures(lSigIndex - 1)
    Else
        Erase tRet.tSignatures
    End If
    
    ExtractModuleDefinition = tRet
    
    Exit Function
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
                       
End Function

Private Function ExtractSignatureDefinition( _
                 ByVal cStm As CROStream) As tSignature
    Const PROC_NAME = "ExtractSignatureDefinition", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lDataSize   As Long
    Dim cSigStream  As CROStream
    Dim eTag        As eSignFileTags
    Dim lTgtIndex   As Long
    Dim tRet        As tSignature
    
    On Error GoTo err_handler
        
    lDataSize = cStm.ReadLong()
    
    Set cSigStream = cStm.ExtractPart(cStm.Position, lDataSize)
    
    cStm.Position = cStm.Position + lDataSize
    
    tRet.sName = cSigStream.ReadMBZStringToUTF16()
    
    Do Until cSigStream.EOF
        
        eTag = cSigStream.ReadByte()
        
        Select Case eTag
        Case TAG_RANGE_LIST
        
            tRet.tRanges = ExtractRangeDefinition(cSigStream)
        
        Case TAG_PATTERN
            
            tRet.bPattern = ExtractPattern(cSigStream)
            
        Case TAG_MASK
            
            tRet.bMask = ExtractMask(cSigStream)
            
        Case TAG_TARGET
            
            If lTgtIndex Then
                If lTgtIndex > UBound(tRet.tTargets) Then
                    ReDim Preserve tRet.tTargets(lTgtIndex + 10)
                End If
            Else
                ReDim tRet.tTargets(9)
            End If
            
            tRet.tTargets(lTgtIndex) = ExtractTargetDefinition(cSigStream)
            
            lTgtIndex = lTgtIndex + 1
            
        Case Else
            ' // Skip
            
            lDataSize = cSigStream.ReadLong()
            cSigStream.Position = cSigStream.Position + lDataSize
            
        End Select
        
    Loop
    
    If lTgtIndex Then
        ReDim Preserve tRet.tTargets(lTgtIndex - 1)
    Else
        Erase tRet.tTargets
    End If
    
    ExtractSignatureDefinition = tRet
    
    Exit Function
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
                       
End Function

Private Function ExtractRangeDefinition( _
                 ByVal cStm As CROStream) As tModuleRange()
    Const PROC_NAME = "ExtractRangeDefinition", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lDataSize   As Long
    Dim cRngStream  As CROStream
    Dim cTmpStream  As CROStream
    Dim eTag        As eSignFileTags
    Dim lRngIndex   As Long
    Dim tRet()      As tModuleRange
    
    On Error GoTo err_handler
        
    lDataSize = cStm.ReadLong()
    
    Set cRngStream = cStm.ExtractPart(cStm.Position, lDataSize)
    
    cStm.Position = cStm.Position + lDataSize
    
    Do Until cRngStream.EOF
        
        eTag = cRngStream.ReadByte()
        
        Select Case eTag
        Case TAG_MODULE_RANGE
            
            If lRngIndex Then
                If lRngIndex > UBound(tRet) Then
                    ReDim tRet(lRngIndex + 10)
                End If
            Else
                ReDim tRet(9)
            End If
            
            lDataSize = cRngStream.ReadLong()
            
            Set cTmpStream = cRngStream.ExtractPart(cRngStream.Position, lDataSize)
            
            tRet(lRngIndex).sSectionName = cTmpStream.ReadMBZStringToUTF16()
            tRet(lRngIndex).lStartOffset = cTmpStream.ReadLong()
            tRet(lRngIndex).lEndOffset = cTmpStream.ReadLong()
            
            lRngIndex = lRngIndex + 1
            
            cRngStream.Position = cRngStream.Position + lDataSize
            
        Case Else
            ' // Skip
            
            lDataSize = cRngStream.ReadLong()
            cRngStream.Position = cRngStream.Position + lDataSize
            
        End Select
        
    Loop
    
    If lRngIndex Then
        ReDim Preserve tRet(lRngIndex - 1)
    Else
        Erase tRet
    End If
    
    ExtractRangeDefinition = tRet
    
    Exit Function
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
                       
End Function

Private Function ExtractTargetDefinition( _
                 ByVal cStm As CROStream) As tSignatureTarget
    Const PROC_NAME = "ExtractTargetDefinition", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lDataSize   As Long
    Dim eType       As eTargetType
    Dim cTgtStream  As CROStream
    Dim tRet        As tSignatureTarget
    
    On Error GoTo err_handler
        
    lDataSize = cStm.ReadLong()
    
    Set cTgtStream = cStm.ExtractPart(cStm.Position, lDataSize)
    
    cStm.Position = cStm.Position + lDataSize
    
    tRet.eType = cTgtStream.ReadByte()
    tRet.sName = cTgtStream.ReadMBZStringToUTF16()
    tRet.lOffset = cTgtStream.ReadLong()

    ExtractTargetDefinition = tRet
    
    Exit Function
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
                       
End Function

Private Function ExtractPattern( _
                 ByVal cStream As CROStream) As Byte()
    Const PROC_NAME = "ExtractPattern", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lDataSize   As Long
    Dim bRet()      As Byte
    
    On Error GoTo err_handler
    
    lDataSize = cStream.ReadLong()
    
    If lDataSize <= 0 Then
        Err.Raise 5
    End If
    
    ReDim bRet(lDataSize - 1)
    
    cStream.ReadData VarPtr(bRet(0)), lDataSize
    
    ExtractPattern = bRet
    
    Exit Function
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
                       
End Function

Private Function ExtractMask( _
                 ByVal cStream As CROStream) As Byte()
    Const PROC_NAME = "ExtractMask", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lDataSize   As Long
    Dim bRet()      As Byte
    
    On Error GoTo err_handler
    
    lDataSize = cStream.ReadLong()
    
    If lDataSize <= 0 Then
        Err.Raise 5
    End If
    
    ReDim bRet(lDataSize - 1)
    
    cStream.ReadData VarPtr(bRet(0)), lDataSize
    
    ExtractMask = bRet
    
    Exit Function
    
err_handler:

    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
                       
End Function

Private Sub Class_Terminate()
    ReleaseSearch
End Sub
