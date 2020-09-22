Attribute VB_Name = "modFormTranslucency"
Option Explicit
'~modFormTranslucency.bas;modGetOSType.bas;modGetTopLevelWindow.bas;
'Windows 2000/XP Form translucency support
'*****************************************************************************
' modFormTranslucency - This modules allows you to set the tranlucency of a
'                       top-level form (one without a parent or owner). If the
'                       specified form is ownder or is a child, its top-level
'                       parent will be modified, and the effects will cascade
'                       through all its children. This modules works only for
'                       Windows 2000/XP. Other OS system will be ignored.
' AlphaBlend range is:
'  0   = fully transparent
'  255 = fully solid
'
'The following fuctions are provided:
' SetWindowTranslucency():      Set the translucency of a form
' SetWindowColorTranslucency(): Set the translucency of a color on a form
' ClearWindowTranslucency():    Remove translucency from a form (Same as sending
'                               an Alpha parameter of 255).
'
' The ColoeKey parameter in SetWindowColorTranslucency allows you to assign an
' RGB color value to the the only color that is modified. For example, if you
' send RGB(0,0,255) or vbBlue, this will set the Alpha value to only colors on
' the form that are Blue.
'
' You can only set translucency on one form at a time. Attempting to perform
' more than one will work only on the first, until translucency is turned off,
'either by calling ClearWindowTranslucency(), or by setting an Alpha value of 255.
'
' NOTE: This modules uses "modGetOSType.bas"
' NOTE: This modules uses "modGetTopLevelWindow.bas"
'*****************************************************************************

'*****************************************************************************
' API calls and constants used
'*****************************************************************************
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const LWA_ALPHA = &H2&
Private Const LWA_COLORKEY = &H1&
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000


Dim MyOS As Integer

'*****************************************************************************
' ClearWindowTranslucency(): Remove translucency from a form
'*****************************************************************************
Public Function ClearWindowTranslucency(ByVal hwnd As Long) As Boolean
  ClearWindowTranslucency = SetWindowTranslucency(hwnd, CByte(255))
End Function

'*****************************************************************************
' SetWindowTranslucency(): Set the translucency of a form
'*****************************************************************************
Public Function SetWindowTranslucency(ByVal hwnd As Long, ByVal Alpha As Byte) As Boolean
  Dim OS As Integer
  Dim nStyle As Long, hndl As Long, Value As Long
  
  If MyOS = 0 Then MyOS = GetOSType()                           'check OS
  If MyOS = 6 Or MyOS > 7 Then                                  'allow only NT2000/XP
    hndl = GetTopLevelWindow(hwnd)                              'get top level window of form
    Value = GetWindowLong(hndl, GWL_EXSTYLE)
    nStyle = Value Or WS_EX_LAYERED                             'set layering
    If SetWindowLong(hndl, GWL_EXSTYLE, nStyle) Then            'if we can set it
      SetWindowTranslucency = SetLayeredWindowAttributes(hndl, 0&, CLng(Alpha), LWA_ALPHA)
      If Alpha = 255 Then Call SetWindowLong(hndl, GWL_EXSTYLE, Value) 'turn off layering
    End If
  End If
End Function

'*****************************************************************************
' SetWindowColorTranslucency(): Set the translucency of a color on a form
'*****************************************************************************
Public Function SetWindowColorTranslucency(ByVal hwnd As Long, ByVal Alpha As Byte, ColorKey As Long) As Boolean
  Dim OS As Integer
  Dim nStyle As Long, hndl As Long, Value As Long
  
  If MyOS = 0 Then MyOS = GetOSType()                           'check OS
  If MyOS = 6 Or MyOS > 7 Then                                  'allow only NT2000/XP
    hndl = GetTopLevelWindow(hwnd)                              'get top level window of form
    Value = GetWindowLong(hndl, GWL_EXSTYLE)
    nStyle = Value Or WS_EX_LAYERED                             'set layering
    If SetWindowLong(hndl, GWL_EXSTYLE, nStyle) Then            'if we can set it
      SetWindowColorTranslucency = SetLayeredWindowAttributes(hndl, ColorKey, CLng(Alpha), LWA_COLORKEY)
      If Alpha = 255 Then Call SetWindowLong(hndl, GWL_EXSTYLE, Value) 'turn off layering
    End If
  End If
End Function

