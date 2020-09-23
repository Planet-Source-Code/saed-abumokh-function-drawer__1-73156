VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PS_NULL = 5

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
Dim bi As BITMAPINFO
Dim dcHandle As Long
Dim hdc As Long, hDIB As Long

Public Enum Qualities
    QualityBlackOnWhite = 1
    QualityWhiteOnBlack = 2
    QualityNearisNeigbour = 3
    QualityHalftoneOrBilinear = 4
End Enum


Const RC_PALETTE As Long = &H100
Const SIZEPALETTE As Long = 104
Const RASTERCAPS As Long = 38

Private Type PALETTEENTRY
peRed As Byte
peGreen As Byte
peBlue As Byte
peFlags As Byte
End Type

Private Type LOGPALETTE
palVersion As Integer
palNumEntries As Integer
palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type

Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type

Private Type PicBmp
Size As Long
Type As Long
hBmp As Long
hPal As Long
Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long


Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Sub Create(ByVal bpp As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Planes As Long = 1)
    Dispose
    With bi.bmiHeader
        .biBitCount = bpp
        .biHeight = Height
        .biWidth = Width
        .biPlanes = Planes
        .biSize = Len(bi.bmiHeader)
        .biCompression = 0
        
    End With
    dcHandle = CreateCompatibleDC(0)
    hDIB = CreateDIBSection(dcHandle, bi, 0, 0, 0, 0)
    SelectObject dcHandle, hDIB
End Sub
Friend Property Get Handle()
    Handle = dcHandle
End Property
Public Sub Dispose()
    DeleteObject hDIB
    DeleteDC dcHandle
    ReleaseDC WindowFromDC(dcHandle), dcHandle
End Sub

Public Sub Clear(ByVal bgColor As Long)
    Dim hPen As Long, hBrush As Long, OldPen As Long, OldBrush As Long
    
    hPen = CreatePen(PS_NULL, 0, 0)
    hBrush = CreateSolidBrush(bgColor)
    
    OldPen = SelectObject(dcHandle, hPen)
    OldBrush = SelectObject(dcHandle, hBrush)
    
    Rectangle dcHandle, 0, 0, Width, Height
    
    SelectObject dcHandle, OldPen
    SelectObject dcHandle, OldBrush
    
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Public Sub GetBitmap(ByVal hBitmap As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)

    Dim hPen As Long, hBrush As Long, OldPen As Long, OldBrush As Long
    
    hPen = CreatePen(PS_NULL, 0, 0)
    hBrush = CreatePatternBrush(hBitmap)
    
    OldPen = SelectObject(dcHandle, hPen)
    OldBrush = SelectObject(dcHandle, hBrush)
    
    Rectangle dcHandle, X, Y, Width, Height
    
    SelectObject dcHandle, OldPen
    SelectObject dcHandle, OldBrush
    
    DeleteObject hPen
    DeleteObject hBrush
    
End Sub

Public Function ConvertToBitmap(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As StdPicture
    Set ConvertToBitmap = hDCToPicture(dcHandle, X, Y, Width, Height)
End Function

Public Sub SetToDC(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Quality As Long, ByVal OperationCode As Long)
    StretchBlt hdc, X, Y, Width, Height, dcHandle, SrcX, SrcY, SrcWidth, SrcHeight, OperationCode
    
End Sub

Public Sub GetFromDC(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal OperationCode As Long)
    StretchBlt dcHandle, 0, 0, Width, Height, hdc, X, Y, Width, Height, OperationCode
    'BitBlt dcHandle, 0, 0, Width, Height, hDC, X, Y, OperationCode
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub


Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim r As Long, pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID
    
    'Fill GUID info
    With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
    End With
    
    'Fill picture info
    With pic
    .Size = Len(pic) ' Length of Structure
    .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
    .hBmp = hBmp ' Handle to bitmap
    .hPal = hPal ' Handle to palette (may be null)
    End With
    
    r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
    
    Set CreateBitmapPicture = IPic
End Function

Private Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, _
    ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, r As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE
    
    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a compatible bitmap
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
    
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        'Number of palette entries
        LogPal.palNumEntries = 256
        'Retrieve the system palette entries
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        'Create the palette
        hPal = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        'Realize the palette
        r = RealizePalette(hDCMemory)
    End If
    
    'Copy the source image to our compatible device context
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    
    'Restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    'Select the palette
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    
    'Delete our memory DC
    r = DeleteDC(hDCMemory)
    
    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)
End Function

Friend Property Get Width() As Long
    Width = bi.bmiHeader.biWidth
End Property

Friend Property Get Height() As Long
    Height = bi.bmiHeader.biHeight
End Property
