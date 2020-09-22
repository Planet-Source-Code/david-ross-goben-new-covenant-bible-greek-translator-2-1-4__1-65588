VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'~clsToolTip.cls;
'Support for easily adding multiline tooltips
'************************************************************************
' clsToolTip.cls -- Support for easily adding multiline tooltips, including
'                   support of forced newline characters such as vbCrLf.
'************************************************************************
'
' Definition: at the top of the form, add:
'   Private MyToolTips As clsTooltip   'MyToolTips can be any name useful to you
' In the Form_Load() event, add:
'   Set MyToolTips = New clsTooltip     'declare object
'   MyToolTips.Create Me              'create object
' In the Form_Unload() event, add:
'   Set MyToolTips = Nothing          'destroy object
'--------------------------------------
' To change the font and pointsize used by the tooltip, do this:
'   MyToolTips.SetFont "Courier New", 12  'set Courier New, 12 point (default Tahoma, 8pt)
'
' To change the tooltip display from the screen width to a defined twip width:
'   MyToolTips.MaxTipWidth = 2880     'set to 2 inches in this example (1440 * 2)
'
' Get tooltip width using:
'   Debug.Print MyToolTips.taxtipwidth
'
' To add a tooltip:
'   MyToolTips.AddTool MyCtrl, "Tooltip text" 'You can add vbCrLf to force multiline text
' If you wish to assign all controls on a form to the list, and use their predefined
' tooltips, use code like this (Note: Controls, such as labels and shapes, which have no
' hWnd value, will be ignored):
'   Dim Ctl As Control
'   For Each Ctl In Form1.Controls
'     MyToolTips.AddTool Ctl, Ctl.ToolTipText  'add a control and its tooltip
'     Ctl.ToolTipText = ""                     'remove original tooltip to avoid double display
'   Next Ctl
'
' Get a control's tooltip using (assuming control has already been added via AddTool):
'   Debug.Print MyToolTips.ToolText(MyCntrl)   'grab tooltip
'
' Assign a new tooltip using:
'   MyToolTips.ToolText(MyCntrl) = "New Tooltip text" & vbCrLf & "Force new line"
'
' To change time tip is shown from default 5 seconds (5000 miliseconds):
'   MyToolTips.DelayTime(ttDelayShow) = 20 * 1000 'set to 20 seconds
'
' Timing Defaults:
' ttDelayInitial = 500  (1/2 sec) time before initial show of tip
' ttDelayShow    = 5000  (5 secs) show time of tip
' ttDelayReshow  = 100 (1/10 sec) if reshow (from other control)
' MaxTipWidth    = Screen.Width   width of screen
' Margin Defaults:
' all Margins    = 0 (Distance between border and text)
'************************************************************************
' NOTE: If a control has a tooltip, then this tooltip will be displayed
'       along with any tooltip you may assign to the control using this
'       class. To effectively use this class and still be able to predefine
'       tooltips at development time, you can assign a control in this fashion:
'    MyToolTip.AddTool Ctrl, Ctrl.ToolTipText  'assign control and tooltip
'    Ctrl.ToolTipText = ""                     'remove normal tooltip
'************************************************************************

'************************************************************************
' API-related constants, types, and methods...
'************************************************************************
Private Const TOOLTIPS_CLASS = "tooltips_class32"
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOPREFIX = &H2
Private Const TTS_BALLOON = &H40

Private Const LF_FACESIZE = 32
Private Const OUT_DEFAULT_PRECIS = 0
Private Const FIXED_PITCH = 1
Private Const FF_MODERN = 48
Private Const LOGPIXELSY = 90
Private Const WM_SETFONT = &H30
Private Const WS_POPUP = &H80000000

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Enum TT_Flags
  TTF_IDISHWND = &H1
  TTF_CENTERTIP = &H2
  TTF_RTLREADING = &H4
  TTF_SUBCLASS = &H10
  TTF_TRACK = &H20
  TTF_ABSOLUTE = &H80
  TTF_TRANSPARENT = &H100
  TTF_DI_SETITEM = &H8000&
End Enum

Private Type TOOLINFO
  cbSize As Long
  uFlags As TT_Flags
  hwnd As Long
  uId As Long
  RECT As RECT
  hinst As Long
  lpszText As String
  lParam As Long
End Type

Public Enum ttDelayTimeConstants
  ttDelayDefault = 0
  ttDelayInitial = 3
  ttDelayShow = 2
  ttDelayReshow = 1
  ttDelayMask = 3
End Enum

Public Enum ttMarginConstants
  ttMarginLeft = 0
  ttMarginTop = 1
  ttMarginRight = 2
  ttMarginBottom = 3
End Enum

Private Const WM_USER = &H400

Private Enum TT_Msgs
  TTM_ACTIVATE = (WM_USER + 1)
  TTM_SETDELAYTIME = (WM_USER + 3)
  TTM_RELAYEVENT = (WM_USER + 7)
  TTM_GETTOOLCOUNT = (WM_USER + 13)
  TTM_WINDOWFROMPOINT = (WM_USER + 16)
  TTM_ADDTOOL = (WM_USER + 4)
  TTM_DELTOOL = (WM_USER + 5)
  TTM_NEWTOOLRECT = (WM_USER + 6)
  TTM_GETTOOLINFO = (WM_USER + 8)
  TTM_SETTOOLINFO = (WM_USER + 9)
  TTM_HITTEST = (WM_USER + 10)
  TTM_GETTEXT = (WM_USER + 11)
  TTM_UPDATETIPTEXT = (WM_USER + 12)
  TTM_ENUMTOOLS = (WM_USER + 14)
  TTM_GETCURRENTTOOL = (WM_USER + 15)
  TTM_TRACKACTIVATE = (WM_USER + 17)
  TTM_TRACKPOSITION = (WM_USER + 18)
  TTM_SETTIPBKCOLOR = (WM_USER + 19)
  TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
  TTM_GETDELAYTIME = (WM_USER + 21)
  TTM_GETTIPBKCOLOR = (WM_USER + 22)
  TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
  TTM_SETMAXTIPWIDTH = (WM_USER + 24)
  TTM_GETMAXTIPWIDTH = (WM_USER + 25)
  TTM_SETMARGIN = (WM_USER + 26)
  TTM_GETMARGIN = (WM_USER + 27)
  TTM_POP = (WM_USER + 28)
  TTM_UPDATE = (WM_USER + 29)
End Enum

Private Declare Sub CopyMemoryByte Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Byte, ByVal Source As String, ByVal Length As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

'************************************************************************
' local storage
'************************************************************************
Private m_hwndTT As Long
Private m_cMaxTip As Long
Private m_hFont As Long
Private m_hDC As Long

'************************************************************************
' shut control down
'************************************************************************
Private Sub Class_Terminate()
  If m_hwndTT Then Call DestroyWindow(m_hwndTT)
End Sub

'************************************************************************
' Create(): Create multiline tooltip instance
' If optional parameter AsBalloon is set to True, the tooltip will be
' displayed as a balloon, with a thorn to the cursor location.
'************************************************************************
Public Function Create(Frm As Form, Optional AsBalloon As Boolean = False) As Boolean
  If (m_hwndTT = 0) Then
    Call InitCommonControls
    
    If AsBalloon Then
      m_hwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, _
                                vbNullString, WS_POPUP Or TTS_NOPREFIX Or TTS_BALLOON, _
                                0, 0, _
                                0, 0, _
                                Frm.hwnd, 0, _
                                App.hInstance, ByVal 0)
    Else
      m_hwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, _
                                vbNullString, TTS_ALWAYSTIP, _
                                0, 0, _
                                0, 0, _
                                Frm.hwnd, 0, _
                                App.hInstance, ByVal 0)
    End If
    Create = CBool(m_hwndTT)
  End If
End Function

'************************************************************************
' SetFont(): Change font for tooltip (Default = Arial 8 point)
'************************************************************************
Public Function SetFont(Optional FontName As String = "Arial", _
                        Optional FontSize As Long = 8, _
                        Optional SetBold As Boolean = False, _
                        Optional SetItalic As Boolean = False) As Boolean
  Dim lFont As LOGFONT
  Dim FntName As String
  Dim NewhFont As Long, FntSize As Long
  
  NewhFont = 0
  If (m_hwndTT) Then
    FntSize = FontSize
    If FntSize < 8 Then FntSize = 8
    If FntSize > 72 Then FntSize = 72
    FntName = Trim$(FontName)
    If Len(FntName) = 0 Then FntName = "Arial"
    m_hDC = GetDC(m_hwndTT)
    With lFont
      .lfHeight = -MulDiv(FntSize, GetDeviceCaps(m_hDC, LOGPIXELSY), 72)
      If SetBold Then
        .lfWeight = 700
      Else
        .lfWeight = 0
      End If
      .lfItalic = SetItalic
      .lfOutPrecision = OUT_DEFAULT_PRECIS
      .lfPitchAndFamily = FIXED_PITCH Or FF_MODERN
      Call CopyMemoryByte(.lfFaceName(1), FntName, CLng(Len(FntName)))
    End With
    
    NewhFont = CreateFontIndirect(lFont)             'create font
    If NewhFont Then
      Call SendMessage(m_hwndTT, WM_SETFONT, NewhFont, 0&)
      If m_hFont Then DeleteObject (m_hFont)
      m_hFont = NewhFont
    End If
  End If
  SetFont = CBool(NewhFont)
End Function

'************************************************************************
' AddTool(): Add a control and tooltip
'************************************************************************
Public Function AddTool(Ctrl As Control, Optional sText As String) As Boolean
  Dim ti As TOOLINFO
  
  If m_hwndTT = 0 Then Exit Function
  If (GetToolInfo(Ctrl.hwnd, ti) = False) Then
    With ti
      .cbSize = Len(ti)
      ' TTF_IDISHWND must be specified to tell the tooltip control
      ' to retrieve the control's rect from it's hWnd specified in uId.
      .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
      .hwnd = Ctrl.Container.hwnd
      .uId = Ctrl.hwnd
      
      If Len(sText) Then
        .lpszText = sText
      Else
        .lpszText = "Tool" & ToolCount + 1
      End If
      ' Maintain the maximun tip text length for GetToolInfo
      m_cMaxTip = Max(m_cMaxTip, Len(.lpszText) + 1)
    End With
    ' Returns 1 on success, 0 on failure
    AddTool = SendMessage(m_hwndTT, TTM_ADDTOOL, 0, ti)
  End If   ' GetToolInfo
End Function

'************************************************************************
' RemoveTool(): Remove a control from the list
'************************************************************************
Public Function RemoveTool(Ctrl As Control) As Boolean
  Dim ti As TOOLINFO

  If m_hwndTT = 0 Then Exit Function
  If GetToolInfo(Ctrl.hwnd, ti) Then
    Call SendMessage(m_hwndTT, TTM_DELTOOL, 0, ti)   ' no rtn val
    RemoveTool = True
  End If
End Function

'************************************************************************
' Get/Set background color for balloon
'************************************************************************
Public Property Get BackColor() As OLE_COLOR
  If m_hwndTT = 0 Then Exit Property
  BackColor = SendMessage(m_hwndTT, TTM_GETTIPBKCOLOR, 0, 0)
End Property

Public Property Let BackColor(Clr As OLE_COLOR)
  If m_hwndTT = 0 Then Exit Property
  Call SendMessage(m_hwndTT, TTM_SETTIPBKCOLOR, Clr, 0)   ' no rtn val
End Property

'************************************************************************
' Get/Set foreground color for balloon
'************************************************************************
Public Property Get ForeColor() As OLE_COLOR
  If m_hwndTT = 0 Then Exit Property
  ForeColor = SendMessage(m_hwndTT, TTM_SETTIPTEXTCOLOR, 0, 0)
End Property

Public Property Let ForeColor(Clr As OLE_COLOR)
  If m_hwndTT = 0 Then Exit Property
  Call SendMessage(m_hwndTT, TTM_SETTIPTEXTCOLOR, Clr, 0)   ' no rtn val
End Property

'************************************************************************
' Get/Set delay-time
'************************************************************************
Public Property Get DelayTime(dwType As ttDelayTimeConstants) As Long
  If m_hwndTT = 0 Then Exit Property
  DelayTime = SendMessage(m_hwndTT, TTM_GETDELAYTIME, (dwType And ttDelayMask), 0&)
End Property

Public Property Let DelayTime(dwType As ttDelayTimeConstants, dwMilliSecs As Long)
  If m_hwndTT = 0 Then Exit Property
  Call SendMessage(m_hwndTT, TTM_SETDELAYTIME, (dwType And ttDelayMask), ByVal dwMilliSecs)  ' no rtn val
End Property

'************************************************************************
' get the hWnd value of the object
'************************************************************************
Public Property Get hwnd() As Long
  hwnd = m_hwndTT
End Property

'************************************************************************
' Get/Set the margin of the tooltip
'************************************************************************
Public Property Get Margin(dwType As ttMarginConstants) As Long
  Dim rc As RECT
  
  If m_hwndTT = 0 Then Exit Property
  Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, rc) ' no rtn val
  Select Case dwType
    Case ttMarginLeft:    Margin = rc.Left * Screen.TwipsPerPixelX
    Case ttMarginTop:     Margin = rc.Top * Screen.TwipsPerPixelX
    Case ttMarginRight:   Margin = rc.Right * Screen.TwipsPerPixelX
    Case ttMarginBottom:  Margin = rc.Bottom * Screen.TwipsPerPixelX
  End Select
End Property

Public Property Let Margin(dwType As ttMarginConstants, cTwips As Long)
  Dim rc As RECT
  
  If m_hwndTT = 0 Then Exit Property
  Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, rc) ' no rtn val
  Select Case dwType
    Case ttMarginLeft:      rc.Left = cTwips / Screen.TwipsPerPixelX
    Case ttMarginTop:      rc.Top = cTwips / Screen.TwipsPerPixelX
    Case ttMarginRight:    rc.Right = cTwips / Screen.TwipsPerPixelX
    Case ttMarginBottom: rc.Bottom = cTwips / Screen.TwipsPerPixelX
  End Select
  Call SendMessage(m_hwndTT, TTM_SETMARGIN, 0, rc)
End Property

'************************************************************************
' Get/Set the maximum tip width in twips
'************************************************************************
Public Property Get MaxTipWidth() As Integer
  If m_hwndTT = 0 Then Exit Property
  MaxTipWidth = LOWORD(SendMessage(m_hwndTT, TTM_GETMAXTIPWIDTH, 0, 0)) * Screen.TwipsPerPixelX
End Property

Public Property Let MaxTipWidth(ByVal cx As Integer)
  If m_hwndTT = 0 Then Exit Property
  If cx < 1 Then cx = Screen.Width
  Call SendMessage(m_hwndTT, TTM_SETMAXTIPWIDTH, 0, ByVal CLng(cx / Screen.TwipsPerPixelX))
End Property

'************************************************************************
' Get the number of objects we have control of
'************************************************************************
Public Property Get ToolCount() As Integer
  If m_hwndTT = 0 Then Exit Property
  ToolCount = SendMessage(m_hwndTT, TTM_GETTOOLCOUNT, 0, 0)
End Property

'************************************************************************
' Get/Set the current tooltip for a control
'************************************************************************
Public Property Get ToolText(Ctrl As Control) As String
  Dim ti As TOOLINFO
  
  If m_hwndTT = 0 Then Exit Property
  If GetToolInfo(Ctrl.hwnd, ti, True) Then ToolText = GetStrFromBufferA(ti.lpszText)
End Property

Public Property Let ToolText(Ctrl As Control, sText As String)
  Dim ti As TOOLINFO
  
  If m_hwndTT = 0 Then Exit Property
  If GetToolInfo(Ctrl.hwnd, ti) Then
    ti.lpszText = sText
    m_cMaxTip = Max(m_cMaxTip, Len(sText) + 1)
    Call SendMessage(m_hwndTT, TTM_UPDATETIPTEXT, 0, ti)   ' no rtn val
  End If
End Property

'************************************************************************
' Return True if the control has an hWnd value
'************************************************************************
Private Function IsWindow(Ctrl As Control) As Boolean
  On Error Resume Next
  IsWindow = CBool(Ctrl.hwnd)
End Function

Private Function GetToolInfo(hwndTool As Long, ti As TOOLINFO, Optional fGetText As Boolean = False) As Boolean
  Dim nItems As Integer
  Dim I As Integer
  
  ti.cbSize = Len(ti)
  If fGetText Then ti.lpszText = String$(m_cMaxTip, 0)
    
  nItems = ToolCount
  For I = 0 To nItems - 1
    If SendMessage(m_hwndTT, TTM_ENUMTOOLS, (I), ti) Then
      If (hwndTool = ti.uId) Then
        GetToolInfo = True
        Exit Function
      End If
    End If
  Next
End Function

'************************************************************************
' Returns the low-order word from the given 32-bit value.
'************************************************************************
Private Function LOWORD(dwValue As Long) As Integer
  MoveMemory LOWORD, dwValue, 2
End Function

'************************************************************************
' Returns the larger of the two passed params
'************************************************************************
Private Function Max(param1 As Long, param2 As Long) As Long
  If param1 > param2 Then Max = param1 Else Max = param2
End Function

'************************************************************************
' Grab the text and strip anything on and after an null character
'************************************************************************
Private Function GetStrFromBufferA(szA As String) As String
  Dim I As Long
  
  I = InStr(1, szA, vbNullChar)
  If I <> 0 Then
    GetStrFromBufferA = Left$(szA, I - 1)
  Else
    GetStrFromBufferA = szA
  End If
End Function

