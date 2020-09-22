Attribute VB_Name = "modSetToolbarBkgColor"
Option Explicit
'~modSetToolbarBkgColor.bas;
'Change the toolbar background color
'************************************************************************************
' modSetToolbarBkgColor - The SetToolbarBkgColor() function changes to toolbars in
'                         the application (not just a single window) to a specified
'                         color value. The Color should be an RGB() value. The function
'                         returns false if it fails. The only way it will fail is if
'                         the specified toolbar does not exist.
'Example:
'  SetToolbarBkgColor Me.MyToolBar, RGB(255, 255, 255) 'set toolbar background to white
'************************************************************************************

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const GCL_HBRBACKGROUND = (-10)

Public Function SetToolbarBkgColor(Tbr As Control, ByVal Clr As Long) As Boolean
  Dim hBrush As Long
  
  hBrush = CreateSolidBrush(Clr)
  If hBrush Then hBrush = SetClassLong(Tbr.hwnd, GCL_HBRBACKGROUND, hBrush)
  If hBrush Then hBrush = DeleteObject(hBrush)
  SetToolbarBkgColor = CBool(hBrush)
End Function
