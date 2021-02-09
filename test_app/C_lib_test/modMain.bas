Attribute VB_Name = "modMain"
' //
' // Using Cdecl functions from tlb and declare
' // by The trick 2021
' //

Option Explicit

Private Declare Function snwprintf1 CDecl Lib "msvcrt" _
                         Alias "_snwprintf" ( _
                         ByVal pszBuffer As Long, _
                         ByVal lCount As Long, _
                         ByVal pszFormat As Long, _
                         ByRef pArg1 As Any) As Long
Private Declare Function snwprintf2 CDecl Lib "msvcrt" _
                         Alias "_snwprintf" ( _
                         ByVal pszBuffer As Long, _
                         ByVal lCount As Long, _
                         ByVal pszFormat As Long, _
                         ByRef pArg1 As Any, _
                         ByRef pArg2 As Any) As Long
Private Declare Function wtoi64 CDecl Lib "msvcrt" _
                         Alias "_wtoi64" ( _
                         ByVal psz As Long) As Currency
                         
Sub Main()
    Dim cObj    As IUnknown
    Dim z()     As Long
    Dim sBuf    As String
    
    ' // Check tlb-calls
    Debug.Print SumByte(1, 2)
    Debug.Print SumShort(3, 4)
    Debug.Print SumInt(5, 6)
    Debug.Print SumCur(7, 8)
    Debug.Print SumFlt(9, 10)
    Debug.Print SumDbl(11, 12)
    Debug.Print SumStr(13, 14)
    Debug.Print SumVnt(15, 16)

    Set cObj = CreateStream(ByVal 0&, 0)
   
    z = GetSA(4, 5)
    
    ' // Check declare calls
    sBuf = Space$(255)
    
    Debug.Print Left$(sBuf, snwprintf1(StrPtr(sBuf), Len(sBuf), StrPtr("Test %ld"), ByVal 123&))
    
    Debug.Print Left$(sBuf, snwprintf2(StrPtr(sBuf), Len(sBuf), StrPtr("Test %ld, %s"), ByVal 123&, ByVal StrPtr("Hello")))
    
    Debug.Print wtoi64(StrPtr("1234567"))
    
End Sub
