Attribute VB_Name = "modGetTopLevelWindow"
Option Explicit
'~modGetTopLevelWindow.bas;
'Returns the window handle of the top-level owner or parent of a given form
'*****************************************************************************
' modGetTopLevelWindow - The GetTopLevelWindow() function returns the window
'                        handle of the top-level owner or parent of a given form,
'                        if any. If none, the form's window handle is returned.
'*****************************************************************************
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Function GetTopLevelWindow(ByVal hChild As Long) As Long
  Dim hWnd As Long
  
  hWnd = hChild                               'begin with current window's handle
  Do While IsWindowVisible(GetParent(hWnd))   'while it has a visible parent
    hWnd = GetParent(hWnd)                    'grab the parent
  Loop
  GetTopLevelWindow = hWnd                    'return the resulting window handle
End Function

