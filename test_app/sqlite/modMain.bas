Attribute VB_Name = "modMain"
' //
' // Using sqlite3.dll
' // by The trick 2021
' //

Option Explicit

Private Const SQLITE_OK     As Long = 0
Private Const SQLITE_ROW    As Long = 100

Private Declare Function sqlite3_open CDecl Lib "sqlite3" ( _
                         ByVal filename As String, _
                         ByRef ppDB As OLE_HANDLE) As Long
Private Declare Function sqlite3_prepare_v2 CDecl Lib "sqlite3" ( _
                         ByVal db As OLE_HANDLE, _
                         ByVal zSql As String, _
                         ByVal nByte As Long, _
                         ByRef ppStmt As OLE_HANDLE, _
                         ByRef pzTail As Any) As Long
Private Declare Function sqlite3_step CDecl Lib "sqlite3" ( _
                         ByVal pStmt As OLE_HANDLE) As Long
Private Declare Function sqlite3_finalize CDecl Lib "sqlite3" ( _
                         ByVal pStmt As OLE_HANDLE) As Long
Private Declare Function sqlite3_close CDecl Lib "sqlite3" ( _
                         ByVal ppDB As OLE_HANDLE) As Long
Private Declare Function sqlite3_column_text16 CDecl Lib "sqlite3" ( _
                         ByVal pStmt As OLE_HANDLE, _
                         ByVal iCol As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32" ( _
                         ByRef pOlechar As Any) As Long
Private Declare Function PutMem4 Lib "msvbvm60.dll" ( _
                         ByRef pDst As Any, _
                         ByVal lVal As Long) As Long
                         
Sub Main()
    Dim pDB         As OLE_HANDLE
    Dim pStmt       As OLE_HANDLE
    Dim lResult     As Long
    Dim sBstrRes    As String
    
    lResult = sqlite3_open(":memory:", pDB)
    
    If lResult <> SQLITE_OK Then
        MsgBox "Cannot open database", vbCritical
        GoTo CleanUp
    End If
    
    lResult = sqlite3_prepare_v2(pDB, "SELECT SQLITE_VERSION()", -1, pStmt, ByVal 0&)
    
    If lResult <> SQLITE_OK Then
        MsgBox "Cannot open database", vbCritical
        GoTo CleanUp
    End If
    
    lResult = sqlite3_step(pStmt)
    
    If lResult = SQLITE_ROW Then
    
        PutMem4 ByVal VarPtr(sBstrRes), SysAllocString(ByVal sqlite3_column_text16(pStmt, 0))
        
        Debug.Print sBstrRes
        
    End If
    
CleanUp:
    
    If pStmt Then sqlite3_finalize pStmt
    If pDB Then sqlite3_close pDB
    
End Sub
