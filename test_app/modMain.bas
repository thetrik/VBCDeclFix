Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    Dim cObj As IUnknown
    Dim z() As Long
    
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
    
    
End Sub
