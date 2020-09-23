Attribute VB_Name = "Functions"
Option Explicit


Public Const EncryptionCode As Byte = 130
Public Const EncryptionCode2 As Byte = 131

Public Enum PenHatchStyles
    NoHatch = 0
    DottedMore = HatchStyle30Percent 'HatchStyleDarkDownwardDiagonal
    DottedSoft = HatchStyle10Percent 'HatchStyleDiagonalCross
    Dashed = HatchStyleLargeCheckerBoard
End Enum

Public Type RGB
    r As Byte
    g As Byte
    b As Byte
End Type

Public Type ARGB
    b As Byte
    g As Byte
    r As Byte
    A As Byte
End Type


Public Type LineStyle
    BorderWidth As Long
    DrawStyle As Long
    BorderColor As Long
    Visible As Boolean
End Type

Public MainAxisesStyle As LineStyle
Public MainGridStyle As LineStyle
Public MainNumbersFont As New StdFont, MainNumbersColor As Long, MainNumbersVisible As Boolean
Public SaveAxisesStyle As LineStyle, SaveGridStyle As LineStyle, SaveNumbersFont As New StdFont, SaveNumbersColor As Long, SaveNumbersVisible As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Declare Function GetDesktopWindow Lib "user32" () As Long

Dim token As Long

Public Sub GdipStartUp(ByVal token As Long)
    Dim gp As GdiplusStartupInput
    gp.GdiplusVersion = 1
    GdiplusStartup token, gp, ByVal 0
End Sub

Public Sub DrawLine(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Color As ARGB, ByVal DrawWidth As Single, ByVal PenStyle As PenHatchStyles, ByVal AntiAlias As Boolean)
    Dim token As Long
    Dim ColorL As Long

    Dim Graphics As Long, Pen As Long, Brush As Long
    
    GdipCreateFromHDC hdc, Graphics
    
    If AntiAlias Then
        GdipSetSmoothingMode Graphics, SmoothingModeAntiAlias
    Else
        GdipSetSmoothingMode Graphics, SmoothingModeNone
    End If
    
    CopyMemory ColorL, Color, 4
        
    If PenStyle = NoHatch Then
        GdipCreatePen1 ColorL, DrawWidth * 0.75, UnitPixel, Pen
    Else
        GdipCreateHatchBrush PenStyle, ColorL, &HFFFFFF, Brush
        GdipCreatePen2 Brush, DrawWidth * 0.75, UnitPixel, Pen
    End If
    
    GdipDrawLine Graphics, Pen, X1, Y1, X2, Y2
    
    GdipDeletePen Pen
    GdipDeleteGraphics Graphics
    
End Sub

Public Function NewARGB(ByVal A As Byte, r As Byte, g As Byte, b As Byte) As ARGB
    NewARGB.A = A
    NewARGB.r = r
    NewARGB.g = g
    NewARGB.b = b
End Function

Public Function ColorToRGB(ByVal Color As Long) As RGB
    CopyMemory ColorToRGB, Color, 3
End Function

Public Sub DrawAxisesAndGrid(Object As Object, AxisesStyle As LineStyle, GridStyle As LineStyle, NumsFont As StdFont, NumsColor As Long, NumsVisible As Boolean)
    
    Dim X As Single, Y As Single
    Dim XNumOnAxis As Integer, YNumOnAxis As Integer
    
    Dim SaveFont As StdFont, SaveForeColor As Long, SaveDrawWidth As Integer, SaveDrawStyle As Long
    
    Set SaveFont = Object.Font
    SaveForeColor = Object.ForeColor
    SaveDrawWidth = Object.DrawWidth
    SaveDrawStyle = Object.DrawStyle
        
    
    'Grid
    
    If GridStyle.Visible Then
        Object.DrawWidth = GridStyle.BorderWidth
        Object.DrawStyle = GridStyle.DrawStyle
        Object.ForeColor = GridStyle.BorderColor
        
        For X = 0 To Object.Width / 2
            Object.Line (X + Object.Width / 2, 0)-(X + Object.Width / 2, Object.Height)
            Object.Line (Object.Width / 2 - X, 0)-(Object.Width / 2 - X, Object.Height)
        Next
        
        For Y = 0 To Object.Height / 2
            Object.Line (0, Y + Object.Height / 2)-(Object.Width, Y + Object.Height / 2)
            Object.Line (0, Object.Height / 2 - Y)-(Object.Width, Object.Height / 2 - Y)
        Next
        
    End If
    
    'Axises
    If AxisesStyle.Visible Then
    
        Object.DrawWidth = AxisesStyle.BorderWidth
        Object.DrawStyle = AxisesStyle.DrawStyle
        Object.ForeColor = AxisesStyle.BorderColor

        Object.Line (0, Object.Height / 2)-(Object.Width, Object.Height / 2)
        Object.Line (Object.Width / 2, 0)-(Object.Width / 2, Object.Height)
        
    End If
        
        
    'Numbers
    If NumsVisible Then
        
        Set Object.Font = NumsFont
        Object.ForeColor = NumsColor
        
        For X = -1 To Object.Width
            XNumOnAxis = Int(X - Object.Width / 2) + 1
            Object.CurrentX = X + ((Object.Width / 2) - Int(Object.Width / 2)) - Object.TextWidth(str(X)) / 2
            Object.CurrentY = Object.Height / 2
            If XNumOnAxis <> 0 Then Object.Print XNumOnAxis
        Next
        
        For Y = -1 To Object.Height
            YNumOnAxis = Int(Y - Object.Height / 2) + 1
            Object.CurrentX = Object.Width / 2 - Object.Parent.TextWidth("-X ")
            Object.CurrentY = Y + ((Object.Height / 2) - Int(Object.Height / 2)) - Object.TextHeight(str(Y)) / 2
            If YNumOnAxis <> 0 Then Object.Print -YNumOnAxis
        Next
    
    End If
    
    Set Object.Font = SaveFont
    Object.ForeColor = SaveForeColor
    Object.DrawWidth = SaveDrawWidth
    Object.DrawStyle = SaveDrawStyle
    
End Sub

Public Function LinearFormulaToProfessional(ByVal str As String) As String

    LinearFormulaToProfessional = str
    
    'LinearFormulaToProfessional = Replace(LinearFormulaToProfessional, "*", chr$(215))
    LinearFormulaToProfessional = Replace(LinearFormulaToProfessional, "/", Chr$(247))
    
    'LinearFormulaToProfessional = Replace(LinearFormulaToProfessional, "^1", chr$(185))
    LinearFormulaToProfessional = Replace(LinearFormulaToProfessional, "^2", Chr$(178))
    LinearFormulaToProfessional = Replace(LinearFormulaToProfessional, "^3", Chr$(179))
    
    LinearFormulaToProfessional = Replace(LinearFormulaToProfessional, "(1" & Chr$(247) & "4)", Chr$(188))
    LinearFormulaToProfessional = Replace(LinearFormulaToProfessional, "(1" & Chr$(247) & "2)", Chr$(189))
    LinearFormulaToProfessional = Replace(LinearFormulaToProfessional, "(3" & Chr$(247) & "4)", Chr$(190))
End Function

Public Function ProfessionalFormulaToLinear(ByVal str As String) As String
    ProfessionalFormulaToLinear = str
    'ProfessionalFormulaToLinear = Replace(LCase(ProfessionalFormulaToLinear), chr$(215), "*")
    ProfessionalFormulaToLinear = Replace(LCase(ProfessionalFormulaToLinear), Chr$(247), "/")
    
    'ProfessionalFormulaToLinear = Replace(LCase(ProfessionalFormulaToLinear), chr$(185), "^1")
    ProfessionalFormulaToLinear = Replace(LCase(ProfessionalFormulaToLinear), Chr$(178), "^2")
    ProfessionalFormulaToLinear = Replace(LCase(ProfessionalFormulaToLinear), Chr$(179), "^3")
    
    ProfessionalFormulaToLinear = Replace(LCase(ProfessionalFormulaToLinear), Chr$(188), "(1/4)")
    ProfessionalFormulaToLinear = Replace(LCase(ProfessionalFormulaToLinear), Chr$(189), "(1/2)")
    ProfessionalFormulaToLinear = Replace(LCase(ProfessionalFormulaToLinear), Chr$(190), "(3/4)")
End Function

Public Sub DrawBorder(Object As Object)
    Dim Pixel As Single
    Pixel = Object.Parent.ScaleX(1, vbPixels, Object.ScaleMode)
    Object.Line (0, 0)-(Object.ScaleWidth, 0)
    Object.Line (0, 0)-(0, Object.ScaleHeight)
    Object.Line (0, Object.ScaleHeight - Pixel)-(Object.ScaleWidth, Object.ScaleHeight - Pixel)
    Object.Line (Object.ScaleWidth - Pixel, 0)-(Object.ScaleWidth - Pixel, Object.ScaleHeight)
End Sub

Public Function DrawTangent(Object As Object, ByVal Consts As String, Expression As String, ByVal FunctionX As Double, ByVal FunctNumDigsAfterDecimal As Byte, Width As Single, Height As Single, ByRef Color As ARGB, ByVal DrawWidth As Single, ByVal PenStyle As Long, ByVal AntiAlias As Boolean) As String

    Dim ValueFirst As Double, ValueLast As Double
    Dim xCorrectFirst As Double, yCorrectFirst As Double
    Dim xCorrectLast As Double, yCorrectLast As Double
    
    Dim LinearFunction As String, lfA As Double, lfB As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    
    Dim Pixel As Single
    Pixel = Object.Parent.ScaleX(1, vbPixels, Object.ScaleMode)
        
    Expression = Replace(LCase(Expression), "exp", "nRaisedToE")
    Expression = Replace(LCase(Expression), "log", "logbase")
    Expression = Replace(LCase(Expression), "logbasebase", "logbase")
    Expression = ProfessionalFormulaToLinear(Expression)
    
    GdipStartUp token
    Dim sc2 As New ScriptControl
    sc2.Language = "VBScript"
    sc2.AddCode Consts
    sc2.AddCode LoadEncryptedFile(EncryptionCode, App.path & "\data\more functions.nmf")
    
    sc2.AddCode "function f(x)" & vbNewLine & "f=" & Expression & vbNewLine & "end function"
        
    X1 = FunctionX - 0.0001
    X2 = FunctionX + 0.0001
    Y1 = sc2.Run("f", X1 - Width / 2)
    Y2 = sc2.Run("f", X2 - Width / 2)
    lfA = (Y2 - Y1) / (X2 - X1)
    lfB = (-lfA * X1) + Y1
    
    
        ValueFirst = lfB
        ValueLast = lfA * Width + lfB
                        
        xCorrectFirst = 0
        yCorrectFirst = (Height / 2) - ValueFirst
        xCorrectLast = Width
        yCorrectLast = Height / 2 - ValueLast
        
        
        xCorrectFirst = Object.Parent.ScaleX(xCorrectFirst, Object.ScaleMode, vbPixels)
        yCorrectFirst = Object.Parent.ScaleY(yCorrectFirst, Object.ScaleMode, vbPixels)
        xCorrectLast = Object.Parent.ScaleX(xCorrectLast, Object.ScaleMode, vbPixels)
        yCorrectLast = Object.Parent.ScaleY(yCorrectLast, Object.ScaleMode, vbPixels)
        
        DrawLine Object.hdc, xCorrectFirst, yCorrectFirst, xCorrectLast, yCorrectLast, Color, DrawWidth, PenStyle, AntiAlias
        
    GdiplusShutdown token
    Object.Refresh
    
    
    X1 = X1 - Width / 2
    X2 = X2 - Width / 2
    Y1 = Y1
    Y2 = Y2
    lfA = (Y2 - Y1) / (X2 - X1)
    lfB = (-lfA * X1) + Y1
    
                    
    DrawTangent = CorrectLinearFunction(lfA, lfB, FunctNumDigsAfterDecimal)
End Function
Private Function CorrectLinearFunction(ByVal A As String, ByVal b As Double, ByVal NumDigitsAfterDecimal As Byte) As String

    Dim str As String
    
    If A <> 0 Then str = str & Round(A, NumDigitsAfterDecimal) & "*x"
    If b <> 0 Then str = str & "+" & Round(b, NumDigitsAfterDecimal)
    
    If A = 0 Then
        str = Replace(str, "+", "")
    End If
    If b = 0 Then
        str = Replace(str, "+0", "")
        str = Replace(str, "-0", "")
    End If

    str = Replace(str, "+-", "-")
    If A = 1 Then str = Replace(str, "1*x", "x")
    If A = -1 Then str = Replace(str, "-1*x", "-x")
    
    CorrectLinearFunction = str
End Function
Public Sub DrawFunction(Object As Object, ByVal Consts As String, Expression As String, ByVal X As Single, ByVal Y As Single, Width As Single, Height As Single, ByRef Color As ARGB, ByVal DrawWidth As Single, ByVal PenStyle As Long, ByVal AntiAlias As Boolean)
    
    Dim StrErr As String
    
    Dim Value1 As Double, Value2 As Double
    Dim xCorrect1 As Single, yCorrect1 As Single
    Dim xCorrect2 As Single, yCorrect2 As Single
    
    Dim XCount As Single
    
    Dim Pixel As Single
    Pixel = Object.Parent.ScaleX(1, vbPixels, Object.ScaleMode)
        
    Expression = Replace(LCase(Expression), "exp", "nRaisedToE")
    Expression = Replace(LCase(Expression), "log", "logbase")
    Expression = Replace(LCase(Expression), "logbasebase", "logbase")
    Expression = ProfessionalFormulaToLinear(Expression)
    
    
    
    GdipStartUp token
    Dim sc2 As New ScriptControl
    sc2.Language = "VBScript"
    sc2.AddCode Consts
    sc2.AddCode LoadEncryptedFile(EncryptionCode, App.path & "\Data\More Functions.nmf")
    
    For XCount = -Width / 2 To Width / 2 Step Pixel
    On Error Resume Next

        sc2.AddCode "Function f(x)" & vbNewLine & _
                    "f=" & Expression & vbNewLine & _
                    "End Function"
        
        Value1 = sc2.Run("f", XCount - Pixel)
        Value2 = sc2.Run("f", XCount)
        
        xCorrect1 = (XCount - Pixel) + (Width / 2)
        yCorrect1 = (Height / 2) - Value1
        
        xCorrect2 = XCount + (Width / 2)
        yCorrect2 = (Height / 2) - Value2
        
        xCorrect1 = Object.Parent.ScaleX(xCorrect1, Object.ScaleMode, vbPixels)
        yCorrect1 = Object.Parent.ScaleX(yCorrect1, Object.ScaleMode, vbPixels)
        xCorrect2 = Object.Parent.ScaleX(xCorrect2, Object.ScaleMode, vbPixels)
        yCorrect2 = Object.Parent.ScaleX(yCorrect2, Object.ScaleMode, vbPixels)
        
        If Err.Number = 0 Then
            If Abs(Value2 - Value1) > Height Then
            Else
                DrawLine Object.hdc, xCorrect1 + X, yCorrect1 + Y, xCorrect2 + X, yCorrect2 + Y, Color, DrawWidth, PenStyle, AntiAlias
            End If
            frmMain.StatusBar.Panels(5).Text = ""
        ElseIf Err.Number = 5 Then  'Invalid prodecure call or argument
        ElseIf Err.Number = 6 Then
            StrErr = "Very big numbers might cause the function isn't drawn successfully"
            frmMain.StatusBar.Panels(5).Text = StrErr
        Else
            StrErr = Err.description
            frmMain.StatusBar.Panels(5).Text = StrErr
        End If
        Err.Clear
        
    Next
    GdiplusShutdown token
End Sub
