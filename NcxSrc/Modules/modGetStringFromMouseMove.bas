Attribute VB_Name = "modGetStringFromMouseMove"
Option Explicit
'~modGetStringFromMouseMove.bas;
'Gets a string from a ListBox control when the mouse moves over it
'*********************************************************************
' modGetStringFromMouseMove:
'
' The GetStringFromMouseMove() function gets a string from a ListBox
' control when the mouse moves over it. Optionally return the listindex
' value if the user needs it.
'
' Possible use: Set data line as a tooltip on a listbox:
'
'EXAMPLE:
'Private Sub ListBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  ListBox1.ToolTipText=GetStringFromMouseMove(ListBox1,X,Y)
'End Sub
'*********************************************************************

'*********************************************************************
' API calls used by GetStringFromMouseMove
'*********************************************************************
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LB_ITEMFROMPOINT = &H1A9

'*********************************************************************
' get string from a ListBox control when the mouse moves over it
'*********************************************************************
Public Function GetStringFromMouseMove(Ctrl As Control, X As Single, Y As Single, Optional ListIndex As Variant) As String
'
' present related tip message
'
  Dim lXPoint As Long
  Dim lYPoint As Long
  Dim lIndex As Long

  lXPoint = CLng(X / Screen.TwipsPerPixelX)
  lYPoint = CLng(Y / Screen.TwipsPerPixelY)

  With Ctrl
    ' get selected item from list
    lIndex = SendMessage(.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
    ' return text if data there, or blank if not
    If lIndex < .ListCount Then
      GetStringFromMouseMove = .List(lIndex)
    Else
      GetStringFromMouseMove = vbNullString
    End If
  End With '(Ctrl)
'
'get list index number if the user supplies an Index variable
'
  If Not IsMissing(ListIndex) Then ListIndex = lIndex

End Function

