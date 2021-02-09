Attribute VB_Name = "modMain"
' //
' // qsort C library usage
' // by The trick 2021
' //

Option Explicit

Private Declare Sub qsort CDecl Lib "msvcrt" ( _
                         ByRef pFirst As Any, _
                         ByVal lNumber As Long, _
                         ByVal lSize As Long, _
                         ByVal pfnComparator As Long)
                
Sub Main()
    Dim z() As Long
    Dim i As Long
    Dim s As String
    
    ReDim z(10)
    
    For i = 0 To UBound(z)
        z(i) = Int(Rnd * 1000)
    Next
    
    qsort z(0), UBound(z) + 1, LenB(z(0)), AddressOf Comparator
    
    For i = 0 To UBound(z)
        Debug.Print z(i)
    Next

End Sub

Private Function Comparator CDecl( _
                 ByRef a As Long, _
                 ByRef b As Long) As Long
    Comparator = a - b
End Function


