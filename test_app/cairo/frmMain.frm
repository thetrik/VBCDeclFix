VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cairo test"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   344
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Cairo lib test
' // You can get dll from https://github.com/preshing/cairo-windows/releases
' // It requires VBCdeclFix Add-in installed
' // by The trick 2021
' //

Option Explicit

Private Declare Function cairo_win32_surface_create CDecl Lib "cairo.dll" ( _
                         ByVal hDc As OLE_HANDLE) As OLE_HANDLE
Private Declare Function cairo_create CDecl Lib "cairo.dll" ( _
                         ByVal pSurface As OLE_HANDLE) As OLE_HANDLE
Private Declare Sub cairo_set_line_width CDecl Lib "cairo.dll" ( _
                    ByVal pCr As OLE_HANDLE, _
                    ByVal dValue As Double)
Private Declare Sub cairo_set_source_rgb CDecl Lib "cairo.dll" ( _
                    ByVal pCr As OLE_HANDLE, _
                    ByVal dR As Double, _
                    ByVal dG As Double, _
                    ByVal dB As Double)
Private Declare Sub cairo_rectangle CDecl Lib "cairo.dll" ( _
                    ByVal pCr As OLE_HANDLE, _
                    ByVal dX As Double, _
                    ByVal dY As Double, _
                    ByVal dW As Double, _
                    ByVal dH As Double)
Private Declare Sub cairo_stroke CDecl Lib "cairo.dll" ( _
                    ByVal pCr As OLE_HANDLE)
Private Declare Sub cairo_destroy CDecl Lib "cairo.dll" ( _
                    ByVal pCr As OLE_HANDLE)
Private Declare Sub cairo_surface_destroy CDecl Lib "cairo.dll" ( _
                    ByVal pSurface As OLE_HANDLE)
                    
Private Sub Form_Load()
    Dim pSurf   As Long
    Dim pCr     As Long

    pSurf = cairo_win32_surface_create(Me.hDc)
    pCr = cairo_create(pSurf)
    
    cairo_set_line_width pCr, 3
    cairo_set_source_rgb pCr, 1, 0.5, 0.5
    cairo_rectangle pCr, 10, 10, 300, 200
    cairo_stroke pCr
    
    cairo_destroy pCr
    cairo_surface_destroy pSurf
    
End Sub
