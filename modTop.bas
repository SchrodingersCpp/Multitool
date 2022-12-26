Attribute VB_Name = "modTop"
Option Private Module

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, _
  ByVal cx As Long, ByVal cy As Long, _
  ByVal uFlags As Long) As Long
Private Declare Function GetWindowLongA Lib "user32" _
  (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
  (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const GWL_EXSTYLE = (-20)

Public hWnd As Long

Public Sub SwitchOnTop()
  Dim varTop As Long
  Dim title As String
  Dim t0 As Double
  t0 = Timer
  Do While Timer - t0 < 3
    DoEvents
  Loop
  hWnd = GetForegroundWindow
  title = Space(255)
  GetWindowText hWnd, title, 255
  varTop = GetWindowLongA(hWnd, GWL_EXSTYLE)
  If Left(Right(WorksheetFunction.Base(varTop, 2), 4), 1) = 0 Then
    varTop = HWND_TOPMOST
    Debug.Print hWnd; "topmost "; title
  Else
    varTop = HWND_NOTOPMOST
    Debug.Print hWnd; "non-topmost "; title
  End If
  SetWindowPos hWnd, varTop, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
