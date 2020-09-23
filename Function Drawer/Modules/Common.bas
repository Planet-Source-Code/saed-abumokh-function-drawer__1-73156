Attribute VB_Name = "Common"
Option Explicit

Private Const CLEARTYPE_QUALITY = 6
Private Const NONANTIALIASED_QUALITY = 3
Private Const PROOF_QUALITY = 2

Private Const LOGPIXELSY = 90
Private Const WM_GETFONT = &H31
Private Const WM_SETFONT = &H30
Private Const LF_FACESize = 32
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
        lfFaceName(1 To LF_FACESize) As Byte
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Public Sub SetAntiAliasedLabels(ByVal Form As Form, ByVal OwnerObject As Object)
    Dim Control As Control
    Dim dt As New DrawText
    On Error Resume Next
    For Each Control In Form.Controls
        If TypeOf Control Is Label Then
            If Control.Parent = OwnerObject Then
                dt.Align = Control.Alignment
                dt.SmoothingMode = SmoothingModeClearType
                dt.VerticalAlign = AlignTop
                dt.WordWrap = Control.Parent.hdc
                dt.hdc = OwnerObject.hdc
                dt.MultiLine = True
                Control.Visible = False
                dt.Draw Control.Caption, Control.Left, Control.Top, Control.Width + Control.Left, Control.Height + Control.Top, Control.Font, Control.ForeColor, 0, True, False
            End If
        End If
    Next
End Sub

Public Function HasHWnd(Object As Object) As Boolean
    On Error Resume Next
    Dim H As Long
    H = Object.hWnd
    If Err = 0 Then HasHWnd = True
End Function

Public Sub SetAntiAliasedFontControls(ByVal Object As Object)
    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In Object
        If HasHWnd(ctrl) Then SetAntiAliasedFont ctrl.hWnd
    Next
End Sub

Public Sub SetAntiAliasedFont(ByVal hWnd As Long)
    Dim hFont As Long
    Dim lf As LOGFONT
    hFont = SendMessage(hWnd, WM_GETFONT, 0, ByVal 0)
    GetObject hFont, Len(lf), lf
    lf.lfQuality = CLEARTYPE_QUALITY
    SendMessage hWnd, WM_SETFONT, CreateFontIndirect(lf), True
End Sub

Public Sub SetDefaultQualityFont(ByVal hWnd As Long)
    Dim hFont As Long
    Dim lf As LOGFONT
    hFont = SendMessage(hWnd, WM_GETFONT, 0, ByVal 0)
    GetObject hFont, Len(lf), lf
    lf.lfQuality = PROOF_QUALITY
    SendMessage hWnd, WM_SETFONT, CreateFontIndirect(lf), True
End Sub

Public Function EncryptText(ByVal str As String, EncCode As Byte) As Byte()
    Dim retVal() As Byte
    Dim ch1 As String
    Dim ch2 As String
    Dim i As Integer
    
    ReDim retVal(1 To Len(str))
    For i = 1 To Len(str)
        ch1 = Mid(str, i, Len(str))
        retVal(i) = Asc(ch1) Xor EncCode
    Next
    EncryptText = retVal
End Function

Public Sub SaveEncryptedFile(ByVal str As String, ByVal EncryptionCode As Byte, ByVal FileName As String)
    Dim StrEnc() As Byte
    Dim i As Integer
    On Error Resume Next
    Kill FileName
    StrEnc = EncryptText(str, EncryptionCode)
    Open FileName For Binary As #1
        For i = 1 To UBound(StrEnc)
            Put #1, Seek(1), StrEnc(i)
        Next
    Close #1
End Sub

Public Function LoadEncryptedFile(ByVal EncryptionCode As Byte, ByVal FileName As String) As String
    Dim retVal As String
    Dim StrEnc() As Byte
    Dim i As Integer
    
    ReDim StrEnc(0 To FileLen(FileName))
    Open FileName For Binary As #1
        For i = 1 To UBound(StrEnc)
            Get #1, Seek(1), StrEnc(i)
            StrEnc(i) = StrEnc(i) Xor EncryptionCode
        Next
    Close #1
    
    For i = 1 To UBound(StrEnc)
    retVal = retVal & Chr$(StrEnc(i))
    Next
    LoadEncryptedFile = retVal
End Function

Public Function StringToLineStyle(ByRef str As String) As LineStyle
    Dim StrDetails() As String
    StrDetails = Split(str, ",")
    
    With StringToLineStyle
        .BorderWidth = StrDetails(0)
        .DrawStyle = StrDetails(1)
        .BorderColor = StrDetails(2)
        .Visible = StrDetails(3)
    End With
End Function

Public Sub StringToStdFont(ByRef str As String, FontObject As StdFont)
    Dim StrDetails() As String
    StrDetails = Split(str, ",")
    
    With FontObject
        .Bold = CBool(StrDetails(0))
        .Charset = StrDetails(1)
        .Italic = CBool(StrDetails(2))
        .Name = StrDetails(3)
        .Size = StrDetails(4)
        .Strikethrough = CBool(StrDetails(5))
        .Underline = CBool(StrDetails(6))
        .Weight = StrDetails(7)
    End With
End Sub

Public Function FontColorFromFontOptionsString(ByRef str As String) As Long
    Dim StrDetails() As String
    StrDetails = Split(str, ",")
    FontColorFromFontOptionsString = StrDetails(8)
End Function

Public Function FontVisibleFromFontOptionsString(ByRef str As String) As Long
    Dim StrDetails() As String
    StrDetails = Split(str, ",")
    FontVisibleFromFontOptionsString = CBool(StrDetails(9))
End Function

Public Function LineStyleToString(ByRef ls As LineStyle) As String
    LineStyleToString = CStr(ls.BorderWidth) & "," & CStr(ls.DrawStyle) & "," & CStr(ls.BorderColor) & "," & CStr(ls.Visible)
End Function

Public Function FontOptionsToString(ByVal stdf As StdFont, ByVal FontColor As Long, ByVal FontVisible As Boolean)
    FontOptionsToString = CStr(stdf.Bold) & "," & CStr(stdf.Charset) & "," & CStr(stdf.Italic) & "," & CStr(stdf.Name) & "," & CStr(stdf.Size) & "," & CStr(stdf.Strikethrough) & "," & CStr(stdf.Underline) & "," & CStr(stdf.Weight) & "," & CStr(FontColor) & "," & CStr(FontVisible)
End Function

Public Function GetPicturesFolder() As String
    Dim arr() As String
    Dim MyPicturesFolder As String
    Dim HomeDrive As String
    arr = Split(Environ$(8), "=")
    HomeDrive = Trim(arr(1))
    arr = Split(Environ$(9), "=")
    MyPicturesFolder = HomeDrive & Trim(arr(1)) & "\My Documents\My Pictures"
    GetPicturesFolder = MyPicturesFolder
End Function

Public Function ReadFile(ByVal FileName As String) As String

    Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long

    nFileNum = FreeFile
    
    Open FileName For Input As nFileNum
    lLineCount = 1
    
    Do While Not EOF(nFileNum)
       Line Input #nFileNum, sNextLine
       
       sNextLine = sNextLine & vbCrLf
       sText = sText & sNextLine
    
    Loop
    Close nFileNum
    ReadFile = sText
    
End Function

Public Sub TrimComboItems(ByVal ComboBox As ComboBox)
    On Error Resume Next
    Dim i As Integer
    For i = 0 To ComboBox.ListCount
        If Trim(ComboBox.list(i) = "") Then ComboBox.RemoveItem i
    Next
    If Trim(ComboBox.list(ComboBox.ListCount - 1)) = "" Then ComboBox.RemoveItem ComboBox.ListCount - 1
End Sub

Public Function GetLineFromString(ByVal str As String, ByVal nIndex As Integer) As String
    Dim StrLines() As String
    StrLines = Split(str, vbNewLine)
    GetLineFromString = StrLines(nIndex)
End Function

Public Function GetNumOfLines(ByVal str As String) As Long
    Dim StrLines() As String
    StrLines = Split(str, vbNewLine)
    GetNumOfLines = UBound(StrLines) + 1
End Function


