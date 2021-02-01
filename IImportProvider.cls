VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IImportProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // IImportProvider.cls - interface which provides address of import entry by its name
' // by The trick 2021
' //

Option Explicit

' // Default
Public Property Get Item( _
                    ByRef sName As String) As PTR
Attribute Item.VB_UserMemId = 0
End Property

Public Property Get Exists( _
                    ByRef sName As String) As Boolean
End Property


