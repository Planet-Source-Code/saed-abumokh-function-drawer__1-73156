Attribute VB_Name = "MouseWheel"
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private PrevWndProc As Long
Private Const MK_CONTROL = &H8
Private Const MK_LBUTTON = &H1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = &H2
Private Const MK_SHIFT = &H4

Private Const WHEEL_DELTA = 120

Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_CLOSE = &H10

Private Const GWL_WNDPROC = (-4)

Private Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Button As Integer, Delta As Integer, X As Integer, Y As Integer
    Select Case Msg
    Case Is = WM_MOUSEWHEEL
        Button = LowWord(wParam)
        Delta = HighWord(wParam)
        X = LowWord(lParam)
        Y = HighWord(lParam)
        frmHelp.Mouse_Wheel Delta, Button, X, Y
    Case Is = WM_CLOSE
        SetWindowLong hwnd, GWL_WNDPROC, PrevWndProc
    End Select
    
    WindowProc = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
End Function

Public Sub Hook(ByVal hwnd As Long)
    UnHook hwnd
    PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHook(ByVal hwnd As Long)
    SetWindowLong hwnd, GWL_WNDPROC, PrevWndProc
End Sub

Private Function LowWord(Word As Long)
    LowWord = CInt("&H" & Right$(CompleteZeroes(Hex$(Word), 8), 4))
End Function

Private Function HighWord(Word As Long)
    HighWord = CInt("&H" & Left$(CompleteZeroes(Hex$(Word), 8), 4))
End Function

Private Function CompleteZeroes(ByVal str As String, ByVal NumZeroes As Integer) As String
    Dim RetVal As String
    Dim i As Integer
    If Len(str) > NumZeroes Then
        RetVal = NumZeroes
    Else
        RetVal = String(NumZeroes - Len(str), "0") & str
    End If
    CompleteZeroes = RetVal
End Function
