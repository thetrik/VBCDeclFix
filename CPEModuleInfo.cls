VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPEModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CPEModuleInfo.cls - obtains few information about PE module
' // by The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME       As String = "CPEModuleInfo"

Private m_hModule           As HANDLE
Private m_tDosHeader        As IMAGE_DOS_HEADER
Private m_tFileHeader       As IMAGE_FILE_HEADER
Private m_tOptionalHeader   As IMAGE_OPTIONAL_HEADER
Private m_tSections()       As IMAGE_SECTION_HEADER

Public Sub LoadAndLock( _
           ByRef sName As String)
    Const PROC_NAME = "LoadAndLock", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim hModule         As HANDLE
    Dim tDosHeader      As IMAGE_DOS_HEADER
    Dim tNtHeaders      As IMAGE_NT_HEADERS
    Dim tSections()     As IMAGE_SECTION_HEADER
    Dim lSectionOffset  As Long
    
    On Error GoTo err_handler

    hModule = LoadLibrary(sName)
    
    If hModule = NULL_PTR Then
        Err.Raise 53, , "Module not found"
    End If
    
    memcpy tDosHeader, ByVal hModule, LenB(tDosHeader)
    
    If tDosHeader.e_magic <> IMAGE_DOS_SIGNATURE Then
        Err.Raise 5
    End If
    
    memcpy tNtHeaders, ByVal hModule + tDosHeader.e_lfanew, LenB(tNtHeaders)
    
    ' // Bad but enough here
    If tNtHeaders.Signature <> IMAGE_NT_SIGNATURE Or _
       tNtHeaders.FileHeader.NumberOfSections <= 0 Then
        Err.Raise 5
    End If

    lSectionOffset = tDosHeader.e_lfanew + LenB(tNtHeaders.Signature) + LenB(tNtHeaders.FileHeader) + _
                     tNtHeaders.FileHeader.SizeOfOptionalHeader
    
    ReDim tSections(tNtHeaders.FileHeader.NumberOfSections - 1)
    
    memcpy tSections(0), ByVal hModule + lSectionOffset, LenB(tSections(0)) * tNtHeaders.FileHeader.NumberOfSections
     
    ReleaseModule
    
    m_tDosHeader = tDosHeader
    m_tFileHeader = tNtHeaders.FileHeader
    m_tOptionalHeader = tNtHeaders.OptionalHeader   ' // Bad but enough here
    m_tSections = tSections
    
    m_hModule = hModule
    
    Exit Sub

err_handler:
    
    If hModule Then
        FreeLibrary hModule
    End If
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Source & vbNewLine & Err.Description
    
End Sub

Friend Property Get BaseAddress() As PTR
    Const PROC_NAME = "BaseAddress_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_hModule = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME
    Else
        BaseAddress = m_hModule
    End If
    
End Property

Friend Property Get SectionByName( _
                    ByRef sName As String) As IMAGE_SECTION_HEADER
    Const PROC_NAME = "SectionByName_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bNameMB()   As Byte
    Dim lIndex      As Long
    
    If m_hModule = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    bNameMB = StrConv(sName, vbFromUnicode)
    
    ReDim Preserve bNameMB(7)
    
    For lIndex = 0 To m_tFileHeader.NumberOfSections - 1
        If memcmp(bNameMB(0), m_tSections(lIndex).SectionName(0), 8) = 8 Then
            
            SectionByName = m_tSections(lIndex)
            Exit Property
            
        End If
    Next
    
    Err.Raise 5, FULL_PROC_NAME, "Section not found"
       
End Property

Friend Property Get Section( _
                    ByVal lIndex As Long) As IMAGE_SECTION_HEADER
    Const PROC_NAME = "Section_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_hModule = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME
    ElseIf lIndex < 0 Or lIndex >= m_tFileHeader.NumberOfSections Then
        Err.Raise 9, FULL_PROC_NAME
    End If
    
    Section = m_tSections(lIndex)
    
End Property

Friend Property Get DosHeader() As IMAGE_DOS_HEADER
    Const PROC_NAME = "DosHeader_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_hModule = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    DosHeader = m_tDosHeader
    
End Property

Friend Property Get FileHeader() As IMAGE_FILE_HEADER
    Const PROC_NAME = "FileHeader_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_hModule = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    FileHeader = m_tFileHeader
    
End Property

Friend Property Get OptionalHeader() As IMAGE_OPTIONAL_HEADER
    Const PROC_NAME = "OptionalHeader_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    If m_hModule = NULL_PTR Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    OptionalHeader = m_tOptionalHeader
    
End Property

Private Sub ReleaseModule()
    
    If m_hModule Then
        FreeLibrary m_hModule
    End If
    
    m_hModule = NULL_PTR
    
End Sub

Private Sub Class_Terminate()
    ReleaseModule
End Sub
