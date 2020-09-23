VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TriCenter = 5

Private Const GRADIENT_FILL_RECT_H  As Long = &H0
Private Const GRADIENT_FILL_RECT_V  As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2

Private Enum RGBEnum
    r = 0
    g = 1
    b = 2
End Enum

Enum GradientRectDirection
    horizontal = GRADIENT_FILL_RECT_H
    Vertical = GRADIENT_FILL_RECT_V
    DiagonalFromTop = 3
    DiagonalFromBottom = 4
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   alpha As Integer
End Type

Private Type GRADIENT_TRIANGLE
   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long
End Type

Private Type GRADIENT_RECT
   UpperLeft As Long
   LowerRight As Long
End Type

Private Type PointColor
    X As Long
    Y As Long
    Color As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDestDC As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hSrcDC As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCT As Long) As Long

Private Function Average(ByRef Nums() As Long, NumCount) As Double
    Dim cPlus
    Dim i As Integer
    
    For i = 0 To NumCount - 1
        cPlus = cPlus + Nums(i)
    Next
    Average = cPlus / NumCount
End Function
Private Function FindRed(Color) As Long
    Dim Red, Green, Blue
    Blue = Color \ 65536
    Green = (Color - (Blue * 65536)) \ 256
    Red = Color - ((Blue * 65536) + (Green * 256))
    FindRed = Red
End Function
Private Function FindGreen(Color) As Long
    Dim Green, Blue
    Blue = Color \ 65536
    Green = (Color - (Blue * 65536)) \ 256
    FindGreen = Green
End Function
Private Function FindBlue(Color) As Long
    Dim Blue
    Blue = Color \ 65536
    FindBlue = Blue
End Function

Private Function GetColorLevel(color1 As Long, color2 As Long, ColorLevel As Byte) As Long
    On Error Resume Next
    Dim R1, G1, B1
    Dim R2, G2, B2
    Dim stepRed, stepGreen, stepBlue
    Dim lvlRed, lvlGreen, lvlBlue
    
    R1 = FindRed(color1)
    G1 = FindGreen(color1)
    B1 = FindBlue(color1)
    
    R2 = FindRed(color2)
    G2 = FindGreen(color2)
    B2 = FindBlue(color2)
        
    stepRed = Abs(R1 - R2) / 255
    stepGreen = Abs(G1 - G2) / 255
    stepBlue = Abs(B1 - B2) / 255
    
    If R1 > R2 Then
        lvlRed = R1 - (ColorLevel * stepRed)
    ElseIf R2 > R1 Then
        lvlRed = (ColorLevel * stepRed) + R1
    ElseIf R1 = R2 Then
        lvlRed = R1
    End If
    
    If G1 > G2 Then
        lvlGreen = G1 - (ColorLevel * stepGreen)
    ElseIf G2 > G1 Then
        lvlGreen = (ColorLevel * stepGreen) + G1
    ElseIf G1 = G2 Then
        lvlGreen = G1
    End If
    
    
    If B1 > B2 Then
        lvlBlue = B1 - (ColorLevel * stepBlue)
    ElseIf B2 > B1 Then
        lvlBlue = (ColorLevel * stepBlue) + B1
    ElseIf B1 = B2 Then
        lvlBlue = B1
    End If
    
    GetColorLevel = RGB(lvlRed, lvlGreen, lvlBlue)
    
End Function

Private Function LongToSignedShort(dwUnsigned As Long) As Integer
    
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
End Function


Private Function RedColor(Color As Long) As Long
    RedColor = LongToSignedShort((Color And &HFF&) * &H100)
End Function

Private Function GreenColor(Color As Long) As Long
    GreenColor = LongToSignedShort(((Color And &HFF00&) \ &H100&) * &H100)
End Function

Private Function BlueColor(Color As Long) As Long
    BlueColor = LongToSignedShort(((Color And &HFF0000) \ &H10000) * &H100)
End Function

Private Sub GradientRectHV(hdc As Long, PointColor1 As PointColor, PointColor2 As PointColor, ByVal Direction As GradientRectDirection, ByVal alpha As Byte)
    Dim dc As New DeviceContext
    dc.Create 24, PointColor2.X - 0, PointColor2.Y - 0
    dc.GetFromDC hdc, PointColor1.X, PointColor1.Y, PointColor2.X, PointColor2.Y, vbSrcCopy
    
    If Direction = horizontal Or Direction = Vertical Then
        Dim RVetrex(0 To 1) As TRIVERTEX
        Dim GradRect As GRADIENT_RECT
        Dim AvgColor As Long
        
        With RVetrex(0)
            .X = 0
            .Y = 0
            .alpha = 0
            .Red = RedColor(PointColor1.Color)
            .Green = GreenColor(PointColor1.Color)
            .Blue = BlueColor(PointColor1.Color)
        End With
        
        With RVetrex(1)
            .X = PointColor2.X - PointColor1.X
            .Y = PointColor2.Y - PointColor1.Y
            .alpha = 0
            .Red = RedColor(PointColor2.Color)
            .Green = GreenColor(PointColor2.Color)
            .Blue = BlueColor(PointColor2.Color)
        End With
        
        GradRect.UpperLeft = 0
        GradRect.LowerRight = 1
        
        GradientFill dc.Handle, RVetrex(0), 2, GradRect, 1, Abs(Direction)
    Else
        If Direction = DiagonalFromBottom Then
            AvgColor = GetColorLevel(PointColor1.Color, PointColor2.Color, 128)
            GradientTriangle dc.Handle, InputPointColor(0, 0, PointColor1.Color), InputPointColor(PointColor2.X - PointColor1.X, 0, AvgColor), InputPointColor(PointColor1.X, PointColor2.Y - PointColor1.Y, AvgColor), alpha
            GradientTriangle dc.Handle, InputPointColor(PointColor2.X - PointColor1.X, 0, AvgColor), InputPointColor(0, PointColor2.Y - PointColor1.Y, AvgColor), InputPointColor(PointColor2.X - PointColor1.X, PointColor2.Y - PointColor1.Y, PointColor2.Color), alpha
        ElseIf Direction = DiagonalFromTop Then
            AvgColor = GetColorLevel(PointColor1.Color, PointColor2.Color, 128)
            GradientTriangle dc.Handle, InputPointColor(0, 0, AvgColor), InputPointColor(PointColor2.X - PointColor1.X, 0, PointColor1.Color), InputPointColor(PointColor2.X - PointColor1.X, PointColor2.Y - PointColor1.Y, AvgColor), alpha
            GradientTriangle dc.Handle, InputPointColor(0, 0, AvgColor), InputPointColor(0, PointColor2.Y - PointColor1.Y, PointColor2.Color), InputPointColor(PointColor2.X - PointColor1.X, PointColor2.Y - PointColor1.Y, AvgColor), alpha
        End If
    End If
    AlphaBlend hdc, PointColor1.X, PointColor1.Y, PointColor2.X, PointColor2.Y, dc.Handle, 0, 0, dc.Width, dc.Height, RGB(0, 0, alpha)
    dc.Dispose
End Sub

Private Sub GradientTriangle(hdc As Long, PointColor1 As PointColor, PointColor2 As PointColor, PointColor3 As PointColor, ByVal alpha As Byte)
    Dim TVetrex(0 To 2) As TRIVERTEX
    Dim pt(0 To 2) As POINTAPI
    Dim rcTri As RECT
    
    Dim GradTriangle As GRADIENT_TRIANGLE
    Dim dc As New DeviceContext
    Dim i As Integer
    
    
    With TVetrex(0)
        .X = PointColor1.X
        .Y = PointColor1.Y
        .alpha = 32767
        .Red = RedColor(PointColor1.Color)
        .Green = GreenColor(PointColor1.Color)
        .Blue = BlueColor(PointColor1.Color)
    End With
    
    With TVetrex(1)
        .X = PointColor2.X
        .Y = PointColor2.Y
        .alpha = 0
        .Red = RedColor(PointColor2.Color)
        .Green = GreenColor(PointColor2.Color)
        .Blue = BlueColor(PointColor2.Color)
    End With
    
    With TVetrex(2)
        .X = PointColor3.X
        .Y = PointColor3.Y
        .alpha = 0
        .Red = RedColor(PointColor3.Color)
        .Green = GreenColor(PointColor3.Color)
        .Blue = BlueColor(PointColor3.Color)
    End With
    
    With GradTriangle
        .Vertex1 = 0
        .Vertex2 = 1
        .Vertex3 = 2
    End With
    
    For i = 0 To 2
        pt(i).X = TVetrex(i).X
        pt(i).Y = TVetrex(i).Y
    Next
    GetRgnBox CreatePolygonRgn(pt(0), 3, 2), rcTri
    
    With rcTri
        
        dc.Create 24, .Right - .Left, .Bottom - .Top
        dc.GetFromDC hdc, .Left, .Top, dc.Width, dc.Height, vbSrcCopy
        
        For i = 0 To 2
            TVetrex(i).X = TVetrex(i).X - .Left
            TVetrex(i).Y = TVetrex(i).Y - .Top
        Next
        
        GradientFill dc.Handle, TVetrex(0), 3, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
        AlphaBlend hdc, .Left, .Top, dc.Width, dc.Height, dc.Handle, 0, 0, dc.Width, dc.Height, RGB(0, 0, alpha)
        dc.Dispose
        
    End With
End Sub

Private Sub GradientLineH(hdc As Long, X1 As Long, X2 As Long, Y As Long, color1 As Long, color2 As Long)
    
    Dim TVetrex(0 To 1) As TRIVERTEX
    Dim GradRect As GRADIENT_RECT
    
    With TVetrex(0)
        .X = X1
        .Y = Y
        .alpha = 0
        .Red = RedColor(color1)
        .Green = GreenColor(color1)
        .Blue = BlueColor(color1)
    End With
    
    With TVetrex(1)
        .X = X2
        .Y = Y + 1
        .alpha = 0
        .Red = RedColor(color2)
        .Green = GreenColor(color2)
        .Blue = BlueColor(color2)
    End With
    
    GradRect.UpperLeft = 0
    GradRect.LowerRight = 1
    
    GradientFill hdc, TVetrex(0), 2, GradRect, 1, GRADIENT_FILL_RECT_H
End Sub

Private Sub GradientRect4Corners(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, color1 As Long, color2 As Long, Color3 As Long, Color4 As Long, ByVal alpha As Byte)
    On Error Resume Next
    
    Dim dc As New DeviceContext
    Dim i As Long
    Dim pColor1 As PointColor, pColor2 As PointColor
    
    dc.Create 24, X2 - X1, Y2 - Y1
    
    For i = Y1 To Y2
        pColor1.X = 0
        pColor1.Y = i - Y1
        pColor1.Color = GetColorLevel(color1, Color3, 256 / (Y2 - Y1) * (i - Y1))
        
        pColor2.X = X2 - X1
        pColor2.Y = i - Y1
        pColor2.Color = GetColorLevel(color2, Color4, 256 / (Y2 - Y1) * (i - Y1))
        
        GradientLineH dc.Handle, pColor1.X, pColor2.X, pColor1.Y, pColor1.Color, pColor2.Color
    Next
    
    AlphaBlend hdc, X1, Y1, dc.Width, dc.Height, dc.Handle, 0, 0, dc.Width, dc.Height, RGB(0, 0, alpha)
    dc.Dispose
End Sub

Private Sub GradientPolyFromCenter(hdc As Long, Xs() As Long, Ys() As Long, Colors() As Long, CenterPointColor As Long, AutoCenterPointColor As Boolean, CenterPointX As Single, AutoCenterPointX As Boolean, CenterPointY As Single, AutoCenterPointY As Boolean, ByVal alpha As Byte)
    
    Dim ptColor() As PointColor
    Dim AvgX
    Dim AvgY
    Dim AvgColors
    Dim nCount
    Dim PointColorCenter As PointColor
    Dim Rs() As Long, Bs() As Long, Gs() As Long
    Dim AvgR, AvgG, AvgB
    Dim PointColorCount As Integer
    
    PointColorCount = UBound(Colors)
    ReDim ptColor(PointColorCount)
    ReDim Rs(PointColorCount)
    ReDim Bs(PointColorCount)
    ReDim Gs(PointColorCount)



    For nCount = 0 To PointColorCount

        ptColor(nCount).X = Xs(nCount)
        ptColor(nCount).Y = Ys(nCount)
        ptColor(nCount).Color = Colors(nCount)
        
        If AutoCenterPointColor = True Then
            Rs(nCount) = FindRed(Colors(nCount))
            Gs(nCount) = FindGreen(Colors(nCount))
            Bs(nCount) = FindBlue(Colors(nCount))
            
            AvgR = Average(Rs, PointColorCount)
            AvgG = Average(Gs, PointColorCount)
            AvgB = Average(Bs, PointColorCount)
            AvgColors = RGB(AvgR, AvgG, AvgB)
            PointColorCenter.Color = AvgColors
        Else
            PointColorCenter.Color = CenterPointColor
        End If
    Next
    
    If AutoCenterPointX = False Then
        PointColorCenter.X = CenterPointX
    Else
        AvgX = Average(Xs, PointColorCount)
        PointColorCenter.X = AvgX
    End If
    
    If AutoCenterPointY = False Then
        PointColorCenter.Y = CenterPointY
    Else
        AvgY = Average(Ys, PointColorCount)
        PointColorCenter.Y = AvgY
    End If



    For nCount = 0 To PointColorCount - 1
        GradientTriangle hdc, ptColor(nCount), ptColor(nCount + 1), PointColorCenter, alpha
    Next
    GradientTriangle hdc, ptColor(nCount), ptColor(1), PointColorCenter, alpha

End Sub

Public Sub Rectangle(ByVal hdc As Long, X1, Y1, X2, Y2, FromColor, ToColor, ByVal GradientDirection As GradientRectDirection, Optional ByVal alpha As Byte = 255)
    GradientRectHV hdc, _
                   InputPointColor(X1, Y1, FromColor), _
                   InputPointColor(X2, Y2, ToColor), _
                   Abs(GradientDirection), alpha
End Sub

Public Sub Triangle(ByVal hdc As Long, X1, Y1, color1, X2, Y2, color2, X3, Y3, Color3, Optional ByVal alpha As Byte = 255)
    GradientTriangle hdc, _
                     InputPointColor(X1, Y1, color1), _
                     InputPointColor(X2, Y2, color2), _
                     InputPointColor(X3, Y3, Color3), alpha
End Sub

Public Sub Rectangle4Colors(ByVal hdc As Long, X1, Y1, X2, Y2, color1, color2, Color3, Color4, Optional ByVal alpha As Byte = 255)
    GradientRect4Corners hdc, CLng(X1), CLng(Y1), CLng(X2), CLng(Y2), CLng(color1), CLng(color2), CLng(Color3), CLng(Color4), alpha
End Sub

Public Sub PolygonFromCenter(hdc As Long, Xs() As Long, Ys() As Long, Colors() As Long, PointColorCount As Long, CenterPointColor As Long, AutoCenterPointColor As Boolean, CenterPointX As Single, AutoCenterPointX As Boolean, CenterPointY As Single, AutoCenterPointY As Boolean, Optional ByVal alpha As Byte = 255)
    GradientPolyFromCenter hdc, Xs(), Ys(), Colors(), CenterPointColor, AutoCenterPointColor, CenterPointX, AutoCenterPointX, CenterPointY, AutoCenterPointY, alpha
End Sub

Private Function InputPointColor(ByVal X, ByVal Y, ByVal Color As Long) As PointColor
    With InputPointColor
        .X = CLng(X): .Y = CLng(Y): .Color = CLng(Color)
    End With
End Function

Private Function DCWidth(ByVal hdc As Long) As Long
    Dim r As RECT
    GetClientRect WindowFromDC(hdc), r
    DCWidth = r.Right - r.Left
End Function

Private Function DCHeight(ByVal hdc As Long) As Long
    Dim r As RECT
    GetClientRect WindowFromDC(hdc), r
    DCHeight = r.Bottom - r.Top
End Function
