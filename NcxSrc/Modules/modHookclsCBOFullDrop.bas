Attribute VB_Name = "modHookclscboFullDrop"
Option Explicit
'~ModHookclsCBOFullDrop.bas;clsCBOFullDrop.cls;
'Module used by clsCBOFullDrop.cls
' *************************************************************************
' API Stuff
' *************************************************************************
Private Const GWL_WNDPROC As Long = -4&
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
' Property names used to stash info within window props.
'
Private Const NewWndProc As String = "NewWndProc"
Private Const OldWndProc As String = "OldWndProc"

'*******************************************************************************
' Subroutine Name   : HookhWnd
' Purpose           : Hook a procedure
'*******************************************************************************
Public Sub HookhWnd(hwnd As Long, WndProcObj As Object)
  If GetProp(hwnd, OldWndProc) <> 0 Then Exit Sub                   'already hooked
  Call SetProp(hwnd, NewWndProc, ObjPtr(WndProcObj))                'save proc to invoke
  Call SetProp(hwnd, OldWndProc, GetWindowLong(hwnd, GWL_WNDPROC))  'save proc hooking before
  Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MyHook)           'insert our hook
End Sub

'*******************************************************************************
' Subroutine Name   : UnhookhWnd
' Purpose           : Unhook a procedure
'*******************************************************************************
Public Sub UnhookhWnd(hwnd As Long)
  Dim lpWndProc As Long
    
  lpWndProc = GetProp(hwnd, OldWndProc)               'get procedure hooked before
  If (lpWndProc <> 0) Then                            'if defined...
     Call SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc) 'plug back into chain
  End If
  Call RemoveProp(hwnd, NewWndProc)                   'then remove properties...
  Call RemoveProp(hwnd, OldWndProc)
End Sub

'*******************************************************************************
' Function Name     : MyHook
' Purpose           : Invoke hooked procedure
'*******************************************************************************
Public Function MyHook(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
  Dim obj As clscboFullDrop                         'pointer to a clsFullDrop object
  Dim lpObjPtr As Long
  
  lpObjPtr = GetProp(hwnd, NewWndProc)              'get prodedure to invoke
  CopyMemory obj, lpObjPtr, 4                       'copy to object pointer
  MyHook = obj.MyWindowProc(hwnd, msg, wp, lp)      'invoke
  CopyMemory obj, Nothing, 4                        'clear hook
End Function

'*******************************************************************************
' Function Name     : InvokeWindowProc
' Purpose           : Invoke pointed procedure
'*******************************************************************************
Public Function InvokeWindowProc(hwnd As Long, msg As Long, wp As Long, lp As Long) As Long
   InvokeWindowProc = CallWindowProc(GetProp(hwnd, OldWndProc), hwnd, msg, wp, lp)
End Function

