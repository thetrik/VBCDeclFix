VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPatchEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // IPatchEntry.cls - interface for patch entries
' // by The trick 2021
' //

Option Explicit

' // Initialize patch using specified signatures
Public Sub Initialize( _
           ByVal cScanner As CSignaturesScanner)
           
End Sub

' // Patch
Public Sub Patch()

End Sub

' // Restore
Public Sub Unpatch()

End Sub

