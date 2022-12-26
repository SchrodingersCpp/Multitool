VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufCalc 
   Caption         =   "Multitool"
   ClientHeight    =   405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10575
   OleObjectBlob   =   "ufCalc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1

Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hWnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal X As Long, _
  ByVal Y As Long, _
  ByVal cx As Long, _
  ByVal cy As Long, _
  ByVal uFlags As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public hWnd As Long
Private formula As String

Private Sub btnTop_Click()
  modMain.Top_Click
End Sub

Private Sub UserForm_Initialize()
  Dim i As Byte
  Dim ret As Long
  Dim formHWnd As Long
  Dim bytOpacity As Byte
  Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"
  With cmbPrec
    For i = 0 To 9
      .AddItem i
    Next i
    .Value = 3
  End With
  cbTrailZeros.Value = True
  cbPlus.Value = False
  ufCalc.Show 0
  formHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, Me.Caption)
  If formHWnd = 0 Then
    Debug.Print Err.LastDllError
  End If
  ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
  If ret = 0 Then
    Debug.Print Err.LastDllError
  End If
  bytOpacity = 0
  hWnd = Application.hWnd
  Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
  Call SetLayeredWindowAttributes(Me.hWnd, 0, bytOpacity, LWA_ALPHA)
End Sub

Public Sub ShowCalc()
End Sub

Private Sub txtIn_Change()
  Const title As String = "Multitool"
  Dim outFull As String
  Dim out As String
  Dim fmt As String
  Dim dec As String
  Dim sign As String
  Dim trZeros As String
  sign = ""
  formula = txtIn.Value
  Call AddBracketsToFunction
  Call SymReplace(" ", "") ' space
  Call SymReplace(ChrW(&H3C0), "PI()")  ' pi
  Call SymReplace(ChrW(&HD7), "*") ' x
  Call SymReplace(ChrW(&H221A), "SQRT") ' sqrt
  Call SymReplace(ChrW(&H3016), "(") ' (
  Call SymReplace(ChrW(&H3017), ")") ' )
  Call SymReplace("[", "(") ' (
  Call SymReplace("]", ")") ' )
  Call SymReplace("{", "(") ' (
  Call SymReplace("}", ")") ' )
  Call SymReplace(ChrW(&H2010), "-") ' hyphen
  Call SymReplace(ChrW(&H2012), "-") ' figure dash
  Call SymReplace(ChrW(&H2013), "-") ' en dash
  Call SymReplace(ChrW(&H2014), "-") ' em dash
  Call SymReplace(ChrW(&H2015), "-") ' horizontal bar
  Call SymReplace(ChrW(&H2061), "") ' function application
  Call SymReplace(ChrW(&HB0), "*PI()/180")   ' degree
  Call SymReplace(",", "")   ' comma
  Call SymReplace(ChrW(&H2044), "/")   ' divide
  Call SymReplace(ChrW(&H2308), "ROUNDUP(")   ' round up
  Call SymReplace(ChrW(&H2309), ";0)")   ' round up
  Call SymReplace(ChrW(&H230A), "ROUNDDOWN(")   ' round down
  Call SymReplace(ChrW(&H230B), ";0)")   ' round down
  Call AbsReplace(True)
  If txtIn.Value <> formula Then
    txtIn.Value = formula
  End If
  ' semicolon used as function argument separator
  formula = Replace(txtIn.Value, ";", _
    Application.International(xlListSeparator))
  If IsError(Evaluate("(" & formula & ")")) Then
    out = "..."
    outFull = out
  ElseIf Evaluate("(" & formula & ")") = "" Then
    out = "..."
    outFull = out
  Else
    outFull = Evaluate("(" & formula & ")")
    If IsNumeric(outFull) Then
      If outFull < 0 Then sign = "-"
      out = Round(outFull, cmbPrec.Value)
      outFull = Abs(outFull)
      If (outFull >= 1000) And _
        (InStr(1, outFull, "E", vbTextCompare) = 0) Then
        outFull = Format(CStr(Int(outFull)), "0,000") & _
          Mid(outFull, Len(CStr(Int(outFull))) + 1)
      End If
      If cbTrailZeros.Value Then
        trZeros = "0"
      Else
        trZeros = "#"
      End If
      out = Format(out, "#,##0." & String(cmbPrec.Value, trZeros))
      If Right(out, 1) = "." Then out = Left(out, Len(out) - 1)
      If cbPlus.Value Then
        If out > 0 Then out = "+" & out
      End If
    Else
      out = outFull
    End If
  End If
  txtOut.Value = out
  ufCalc.Caption = title & String(141, " ") & sign & outFull
End Sub

' add brackets to function with single value
Private Sub AddBracketsToFunction()
  Dim i As Integer
  Dim ch As String
  Const n As Integer = 10000 ' arbitrary number of characters
  Dim arrFuncChar()
  arrFuncChar = Array(ChrW(&H2061), ChrW(&H221A))
  Dim arrLbr()
  arrLbr = Array(ChrW(&H3016), "(")
  Dim arrStopChars()
  arrStopChars = Array(" ", "+", "-", "*", "/", ChrW(&HD7), _
    ChrW(&H2010), ChrW(&H2012), ChrW(&H2013), _
    ChrW(&H2014), ChrW(&H2015))
  formula = formula & " "
  For i = Len(formula) - 1 To 1 Step -1
    ch = Mid(formula, i, 1)
    If Not IsError(Application.Match(ch, arrFuncChar, 0)) Then
      ch = Mid(formula, i + 1, 1)
      If IsError(Application.Match(ch, arrLbr, 0)) Then
        formula = Left(formula, i) & "(" & Mid(formula, i + 1, n)
        For j = i + 1 To Len(formula)
          ch = Mid(formula, j, 1)
          If Not IsError(Application.Match(ch, arrStopChars, 0)) Then
            formula = Left(formula, j - 1) & ")" & Mid(formula, j, n)
            Exit For
          End If
        Next j
      End If
    End If
  Next i
End Sub

' process Abs brackets
Private Sub AbsReplace(ByRef val As Boolean)
  Const Lbr As String = "ABS("
  Const Rbr As String = ")"
  Dim n As Integer
  Dim sym As String
  Dim expr As String
  Dim txtLen As Integer
  Dim rpl As String
  expr = formula
  txtLen = Len(expr)
  For n = txtLen To 1 Step -1
    sym = Mid(expr, n, 1)
    If sym = "|" Then
      If n = 1 Then
        rpl = Lbr
      ElseIf n = txtLen Then
        rpl = Rbr
      Else
        If Mid(expr, n - 1, 1) Like "[(/+/*///-]" Then
          rpl = Lbr
        ElseIf Mid(expr, n + 1, 1) Like "[)/+/*///-]" Then
          rpl = Rbr
        End If
      End If
      expr = WorksheetFunction.Replace(expr, n, 1, rpl)
    End If
  Next n
  formula = expr
End Sub

Private Sub SymReplace(ByVal sym As String, ByVal newSym As String)
  If InStr(1, formula, sym) > 0 Then
    formula = Replace(formula, sym, newSym)
  End If
End Sub

Private Sub cmbPrec_Change()
  Call txtIn_Change
End Sub

Private Sub cmbSep_Change()
  Call txtIn_Change
End Sub

Private Sub cbTrailZeros_Click()
  Call txtIn_Change
End Sub

Private Sub cbPlus_Click()
  Call txtIn_Change
End Sub

Private Sub UserForm_Terminate()
  Dim bytOpacity As Byte
  Unload Me
  bytOpacity = 255
  hWnd = Application.hWnd
  Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
  Call SetLayeredWindowAttributes(Me.hWnd, 0, bytOpacity, LWA_ALPHA)
End Sub
