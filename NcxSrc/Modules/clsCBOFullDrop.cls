VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clscboFullDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'~clsFullDrop.cls;ModHookclsCBOFullDrop.bas;
'allows ComboBox dropdowns to display its ENTIRE list
'*********************************************************************
' clsFullDrop: This class module allows ComboBox dropdowns to display its ENTIRE
'            list, not just a partial. Declare one of these object for each
'            ComboBox you wish to affect.
'USAGE:
'''  In form header:
'  Private cboDropHandler As CFullDrop
'''  In Form_Load() event:
'  Set cboDropHandler = New CFullDrop
'  cboDropHandler.hWnd = Combo1.hWnd
'''  In Form_Unload() event:
'  Set cboDropHandler = Nothing
'
' NOTE: This Class uses the module modHookclsFullDrop.bas
'*********************************************************************

' The WM_CTLCOLORLISTBOX message is sent to the parent window of a list box before
' the system draws the list box. By responding to this message, the parent window
' can set the text and background colors of the list box by using the specified
' display device context handle.
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_LBUTTONUP = &H202

'other Const declarations
Private Const LB_GETCOUNT = &H18B
Private Const LB_GETITEMHEIGHT = &H1A1
Private Const SM_CYBORDER = 6

' API Declarations
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'*******************************************************************************
' Private Member variable to store Hwnd value. Save for unhooking
'*******************************************************************************
Private m_hWnd As Long
Private m_oldTop As Long
Private m_oldHeight As Long

'*******************************************************************************
' Subroutine Name   : Class_Terminate
' Purpose           : When terminating, release anything set
'*******************************************************************************
Private Sub Class_Terminate()
  If m_hWnd <> 0 Then
    Call UnhookhWnd(m_hWnd)   'unhook handle
    m_hWnd = 0                'remove tag to prevent redundancy
  End If
End Sub

'*******************************************************************************
' Let Name          : hWnd
' Purpose           : Assign a new hWnd. If one already set, unhook it first
'*******************************************************************************
Public Property Let hwnd(ByVal NewhWnd As Long)
  If m_hWnd <> 0 Then
    Call UnhookhWnd(m_hWnd)   'unhook any previous hWnd
    m_hWnd = 0                'remove tag to prevent redundancy
  End If
  If IsWindow(NewhWnd) Then   'do only if hWnd is for a Window object
    m_hWnd = NewhWnd          'save hwnd for unhook
    Call HookhWnd(m_hWnd, Me) 'pass handle and ClsFullDrop object
  End If
End Property

'*******************************************************************************
' Function Name     : MyWindowProc
' Purpose           : WndProc function
'*******************************************************************************
Friend Function MyWindowProc(hwnd As Long, msg As Long, wp As Long, lp As Long) As Long
  Dim nHeight As Long, iHeight As Long, nWidth As Long, nCount As Long, ScHeight As Long
  Dim rLst As RECT, rCbo As RECT

  Select Case msg
    Case WM_LBUTTONUP       'when left button up (click), reset dims
      m_oldHeight = 0
      m_oldTop = 0
    
    'when preparing to paint, size the droplist to show as much as possible
    Case WM_CTLCOLORLISTBOX
      iHeight = SendMessage(lp, LB_GETITEMHEIGHT, 0, ByVal 0&)  'get single-item height
      nCount = SendMessage(lp, LB_GETCOUNT, 0, ByVal 0&)        'get # if items in list
      Call GetWindowRect(lp, rLst)                              'set full height
      nHeight = (nCount * iHeight) + (GetSystemMetrics(SM_CYBORDER) * 2)
      Call GetWindowRect(hwnd, rCbo)                            'get cbo posn
      ScHeight = (Screen.Height \ Screen.TwipsPerPixelY)        'get screen height in pixels
      nWidth = rLst.Right - rLst.Left                           'get width of box
      If rCbo.Bottom + nHeight > ScHeight Then                  'list below screen?
        If rCbo.Top - nHeight < 0 Then                          'can move above box?
          nHeight = ((ScHeight - rCbo.Bottom + iHeight - 1) \ iHeight) * iHeight 'no, so shorten
        Else
          rLst.Top = rCbo.Top - nHeight    'flip to above box if we can
        End If
        If m_oldHeight <> nHeight Or m_oldTop <> rLst.Top Then  'update only if different
          Call MoveWindow(lp, rLst.Left, rLst.Top, nWidth, nHeight, 1&) 'size the window
        End If
      Else      'nothing existing, no screen bleedover, so just do it
        Call MoveWindow(lp, rLst.Left, rLst.Top, nWidth, nHeight, 1&) 'size the window
      End If
      m_oldHeight = nHeight                                     'save dims
      m_oldTop = rLst.Top
  End Select
  MyWindowProc = InvokeWindowProc(hwnd, msg, wp, lp)            'invoke regular proc
End Function


