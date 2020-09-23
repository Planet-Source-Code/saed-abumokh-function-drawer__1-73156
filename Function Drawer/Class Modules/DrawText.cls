VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DrawText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_HIDEPREFIX = &H100000
Private Const DT_PREFIXONLY = &H200000
Private Const DT_INTERNAL = &H1000
Private Const DT_EDITCONTROL = &H2000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

Private Const NONANTIALIASED_QUALITIY = 3
Private Const ANTIALIASED_QUALITIY = 5
Private Const CLEARTYPE_QUALITIY = 6

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Enum TextVerticalAlign
    AlignTop = DT_TOP
    AlignVCenter = DT_VCENTER
    AlignBottom = DT_BOTTOM
End Enum

Public Enum TextEllipsis
    NoEllipsis = 0
    EndEllipsis = DT_END_ELLIPSIS
    PathEllpsis = DT_PATH_ELLIPSIS
    WordEllipsis = DT_WORD_ELLIPSIS
End Enum

Public Enum PrefixMode
    ShowPrefix = 0
    HidePrefix = DT_HIDEPREFIX
    NoPrefix = DT_NOPREFIX
    PrefixOnly = DT_PREFIXONLY
End Enum

Public Enum SmoothingModes
    SmoothingModeNonAntiAliased = NONANTIALIASED_QUALITIY
    SmoothingModeAntiAliased = ANTIALIASED_QUALITIY
    SmoothingModeClearType = CLEARTYPE_QUALITIY
End Enum

Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal lfHeight As Long, ByVal lfWidth As Long, ByVal lfEscapement As Long, ByVal lfOrientation As Long, ByVal lfWeight As Long, ByVal lfItalic As Long, ByVal lfUnderline As Long, ByVal lfStrikeOut As Long, ByVal lfCharSet As Long, ByVal lfOutPrecision As Long, ByVal lfClipPrecision As Long, ByVal lfQuality As Long, ByVal lfPitchAndFamily As Long, ByVal lfFaceName As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function DrawTextExA Lib "user32" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function DrawTextExW Lib "user32" (ByVal hdc As Long, ByVal lpsz As Long, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long

Dim dthDC As Long
Dim dtSmoothingMode As SmoothingModes
Dim dtMeasureInPexils As Boolean
Dim dtAlign As AlignmentConstants
Dim dtVerticalAlign As TextVerticalAlign
Dim dtEllipsis As TextEllipsis
Dim dtPrefix As PrefixMode
Dim dtWordWrap As Boolean
Dim dtTabStop As Boolean
Dim dtSingleLine As Boolean
Dim dtUseTabWidth As Boolean
Dim dtNoClip As Boolean
Dim dtMultiLine As Boolean
Dim dtRightToLeft As Boolean
Dim dtTabWidth As Long
Dim dtLeftMargin As Long
Dim dtRightMargin As Long

Friend Property Get hdc() As Long
    hdc = dthDC
End Property

Friend Property Let hdc(ByVal vNewValue As Long)
   dthDC = vNewValue
End Property

Friend Property Get SmoothingMode() As SmoothingModes
    SmoothingMode = dtSmoothingMode
End Property

Friend Property Let SmoothingMode(ByVal vNewValue As SmoothingModes)
    dtSmoothingMode = vNewValue
End Property

Friend Property Get MeasureInPixels() As Boolean
    MeasureInPexils = dtMeasureInPexils
End Property

Friend Property Let MeasureInPexils(ByVal vNewValue As Boolean)
   dtMeasureInPexils = vNewValue
End Property

Friend Property Get Align() As AlignmentConstants
    Align = dtAlign
End Property

Friend Property Let Align(ByVal vNewValue As AlignmentConstants)
   dtAlign = vNewValue
End Property

Friend Property Get VerticalAlign() As TextVerticalAlign
    VerticalAlign = dtVerticalAlign
End Property

Friend Property Let VerticalAlign(ByVal vNewValue As TextVerticalAlign)
   dtVerticalAlign = vNewValue
End Property

Friend Property Get Ellipsis() As TextEllipsis
    Ellipsis = dtEllipsis
End Property

Friend Property Let Ellipsis(ByVal vNewValue As TextEllipsis)
   dtEllipsis = vNewValue
End Property

Friend Property Get Prefix() As PrefixMode
    Prefix = dtPrefix
End Property

Friend Property Let Prefix(ByVal vNewValue As PrefixMode)
   dtPrefix = vNewValue
End Property

Friend Property Get WordWrap() As Boolean
    WordWrap = dtWordWrap
End Property

Friend Property Let WordWrap(ByVal vNewValue As Boolean)
   dtWordWrap = vNewValue
End Property

Friend Property Get TabStop() As Boolean
    TabStop = dtTabStop
End Property

Friend Property Let TabStop(ByVal vNewValue As Boolean)
   dtTabStop = vNewValue
End Property

Friend Property Get SingleLine() As Boolean
    SingleLine = dtSingleLine
End Property

Friend Property Let SingleLine(ByVal vNewValue As Boolean)
   dtSingleLine = vNewValue
End Property

Friend Property Get UseTabWidth() As Boolean
    UseTabWidth = dtUseTabWidth
End Property

Friend Property Let UseTabWidth(ByVal vNewValue As Boolean)
   dtUseTabWidth = vNewValue
End Property

Friend Property Get NoClip() As Boolean
    NoClip = dtNoClip
End Property

Friend Property Let NoClip(ByVal vNewValue As Boolean)
   dtNoClip = vNewValue
End Property

Friend Property Get MultiLine() As Boolean
    MultiLine = dtMultiLine
End Property

Friend Property Let MultiLine(ByVal vNewValue As Boolean)
   dtMultiLine = vNewValue
End Property

Friend Property Get RightToLeft() As Boolean
    RightToLeft = dtRightToLeft
End Property

Friend Property Let RightToLeft(ByVal vNewValue As Boolean)
   dtRightToLeft = vNewValue
End Property

Friend Property Get TabWidth() As Long
    TabWidth = dtTabWidth
End Property

Friend Property Let TabWidth(ByVal vNewValue As Long)
   dtTabWidth = vNewValue
End Property

Friend Property Get LeftMargin() As Long
    LeftMargin = dtLeftMargin
End Property

Friend Property Let LeftMargin(ByVal vNewValue As Long)
   dtLeftMargin = vNewValue
End Property

Friend Property Get RightMargin() As Long
    RightMargin = dtRightMargin
End Property

Friend Property Let RightMargin(ByVal vNewValue As Long)
   dtRightMargin = vNewValue
End Property

Public Function Draw(ByVal str As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Font As StdFont, ByVal FontColor As Long, ByVal FontAngle As Single, ByVal FontTransparent As Boolean, ByVal IsUnicode As Boolean)
    Dim dtp As DRAWTEXTPARAMS
    Dim SaveTextColor As Long, SaveBkMode As Long
    Dim r As RECT
    Dim hFont As Long, OldFont As Long
    Dim dtFontSize As Long
    Dim dtFlags As Long
    
    If dtMeasureInPexils Then
        dtFontSize = Font.Size
    Else
        dtFontSize = -MulDiv(Font.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    End If
    
    SaveTextColor = GetTextColor(dthDC)
    SaveBkMode = GetBkMode(dthDC)
    
    SetTextColor dthDC, FontColor
    SetBkMode dthDC, Abs(Not (FontTransparent)) + 1
    
    r.Left = X
    r.Top = Y
    r.Right = Width
    r.Bottom = Height
    
    dtp.cbSize = Len(dtp)
    dtp.iLeftMargin = dtLeftMargin
    dtp.iRightMargin = dtRightMargin
    dtp.iTabLength = dtTabWidth
    
    If dtAlign = vbLeftJustify Then
        dtAlign = DT_LEFT
    ElseIf dtAlign = vbRightJustify Then
        dtAlign = DT_RIGHT
    ElseIf dtAlign = vbCenter Then
        dtAlign = DT_CENTER
    End If
    dtFlags = dtAlign Or dtVerticalAlign Or dtEllipsis Or dtPrefix Or _
              Abs(dtWordWrap * DT_WORDBREAK) Or _
              Abs(dtSingleLine * DT_SINGLELINE) Or _
              Abs(dtUseTabWidth * DT_EXPANDTABS) Or _
              Abs(dtNoClip * DT_WORDBREAK) Or _
              Abs(dtMultiLine * DT_WORDBREAK) Or _
              Abs(dtTabStop * DT_TABSTOP) Or _
              Abs(dtMultiLine * DT_EDITCONTROL) Or _
              Abs(dtRightToLeft * DT_WORDBREAK)
              
    hFont = CreateFont(dtFontSize, 0, FontAngle, 0, Font.Weight, Font.Italic, Font.Underline, Font.Strikethrough, Font.Charset, 0, 0, dtSmoothingMode, 0, Font.Name)
    OldFont = SelectObject(dthDC, hFont)
    
    If IsUnicode Then
        DrawTextExW dthDC, StrPtr(str), Len(str), r, dtFlags, dtp
    Else
        DrawTextExA dthDC, str, Len(str), r, dtFlags, dtp
    End If
    
    SelectObject dthDC, OldFont
    DeleteObject hFont
    
    SetBkMode dthDC, SaveBkMode
    SetTextColor dthDC, SaveTextColor

End Function
