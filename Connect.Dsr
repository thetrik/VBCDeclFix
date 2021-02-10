VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9945
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   17542
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "CDeclFix"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' //
' // Connect.dsr
' // by The trick 2021
' //

Option Explicit

Private Const MODULE_NAME   As String = "Connect"

Private m_cVBInstance   As VBIDE.VBE
Private m_cPatcher      As CRuntimePatcher
Private m_cMenuCmd      As CommandBarControl

Private WithEvents m_cMnuHandler    As CommandBarEvents
Attribute m_cMnuHandler.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection( _
            ByVal Application As Object, _
            ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
            ByVal AddInInst As Object, _
            ByRef custom() As Variant)
    Const PROC_NAME = "AddinInstance_OnConnection", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    On Error GoTo error_handler

    Set m_cVBInstance = Application
    
    Set m_cMenuCmd = AddToAddInCommandBar("CDeclFix")
    
    If Not m_cMenuCmd Is Nothing Then
        
        Set m_cMnuHandler = m_cVBInstance.Events.CommandBarEvents(m_cMenuCmd)
        
    End If
    
    Set m_cPatcher = New CRuntimePatcher

    m_cPatcher.Initialize
    m_cPatcher.CDeclEnabled = True
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Number & vbNewLine & Err.Source & vbNewLine & Err.Description, vbCritical
    
End Sub

Private Sub AddinInstance_OnDisconnection( _
            ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
            custom() As Variant)
    
    Set m_cPatcher = Nothing
    
    If Not m_cMenuCmd Is Nothing Then
        m_cMenuCmd.Delete
    End If
    
End Sub

Private Function AddToAddInCommandBar( _
                 ByRef sCaption As String) As CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu           As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = m_cVBInstance.CommandBars("Add-Ins")
    
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

Private Sub m_cMnuHandler_Click( _
            ByVal CommandBarControl As Object, _
            ByRef handled As Boolean, _
            ByRef CancelDefault As Boolean)
    
    If m_cPatcher.CDeclEnabled Then
        frmAbout.lblState.ForeColor = &HFF7070
        frmAbout.lblState.Caption = "cdecl enabled"
    Else
        frmAbout.lblState.ForeColor = &H7070FF
        frmAbout.lblState.Caption = "cdecl disabled (an error occured)"
    End If
    
    frmAbout.Show vbModal
    
End Sub
