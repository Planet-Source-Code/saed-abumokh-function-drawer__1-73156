VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPlusPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type

Private Type PICTDESC
   Size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type

' GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hdc As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal token As Long)

' GDI and GDI+ constants
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2


Private GdipToken       As Long
Private GdipInitialized As Boolean

Private Const GdiplusVersion     As Long = 1
Private Const CP_ACP            As Long = 0


Private Type ImageCodecInfo
   ClassID As GUID
   FormatID As GUID
   CodecName As Long
   DllName As Long
   FormatDescription As Long
   FilenameExtension As Long
   MimeType As Long
   Flags As Long
   Version As Long
   SigCount As Long
   SigSize As Long
   SigPattern As Long
   SigMask As Long
End Type


Private Enum SmoothingMode
    SmoothingModeInvalid = -1&
    SmoothingModeDefault = 0&
    SmoothingModeHighSpeed = 1&
    SmoothingModeHighQuality = 2&
    SmoothingModeNone = 3&
    SmoothingModeAntiAlias8x4 = 4&
    SmoothingModeAntiAlias = SmoothingModeAntiAlias8x4
End Enum

Private Enum GpStatus
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function lStrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Any, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long


Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long


Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numEncoders As Long, ByRef Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, ByRef Encoders As Any) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMode As SmoothingMode) As GpStatus

Private Declare Function GdipSaveGraphics Lib "gdiplus" (ByVal Graphics As Long, ByVal State As Long) As GpStatus

' Initialises GDI Plus
Private Function InitGDIPlus(ByVal token As Long) As Long
    Dim gdipInit As GdiplusStartupInput
    gdipInit.GdiplusVersion = 1
    GdiplusStartup token, gdipInit, ByVal 0&
    InitGDIPlus = token
End Function

' Frees GDI Plus
Private Sub FreeGDIPlus(token As Long)
    GdiplusShutdown token
End Sub

Public Sub GetImageSize(ByVal PicFile As String, ByVal ImgWidth As Long, ByVal ImageHeight As Long)
    Dim Img As Long, ImgW As Long, ImgH As Long
    GdipLoadImageFromFile StrPtr(PicFile), Img
    GdipGetImageWidth Img, ImgW
    GdipGetImageHeight Img, ImgH
    ImgWidth = ImgW
    ImageHeight = ImgH
    GdipDisposeImage Img
End Sub

'Private Function GetImageSize(ByVal ImageFile As String) As SizeL
'    Dim token As Long
'    Dim Img As Long
'    Dim gps As GdiplusStartupInput
'    gps.GdiplusVersion = 1
'
'    GdiplusStartup token, gps
'
'    GdipLoadImageFromFile StrPtr(ImageFile), Img
'    GdipGetImageWidth Img, GetImageSize.cx
'    GdipGetImageHeight Img, GetImageSize.cy
'    GdipDisposeImage Img
'
'    GdiplusShutdown token
'End Function

' Resize the picture using GDI plus
Public Function LoadPictureToHDC(strFile As String, hdc As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long) As Long
    On Error Resume Next
    Dim Graphics As Long
    Dim Img As Long
    Dim ImgH As Long, ImgW As Long
    
    GdipLoadImageFromFile StrPtr(strFile), Img
    
    GdipGetImageWidth Img, ImgW
    GdipGetImageHeight Img, ImgH
    
    GdipCreateFromHDC hdc, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
        
    LoadPictureToHDC = GdipDrawImageRectRectI(Graphics, Img, DestX, DestY, DestWidth * (DestWidth / ImgW), DestHeight * (DestHeight / ImgH), SrcX, SrcY, SrcWidth, SrcHeight, UnitPixel, 0, 0, 0)
    
    GdipDisposeImage Img
    GdipDeleteGraphics Graphics
End Function

' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub


Private Sub Class_Initialize()
    InitGDIPlus GdipToken
End Sub

Private Sub Class_Terminate()
    FreeGDIPlus GdipToken
End Sub

Public Function SavePictureFromHDC(ByVal hdc As Long, ByVal hBitmap As Long, ByVal sFileName As String) As Boolean
    Dim lBitmap As Long
    Dim PicEncoder As GUID
    Dim sID As String
    Dim Graphics As Long
    Dim State As Long
    
    Select Case LCase$(Right$(sFileName, 4))
        Case ".png"
            sID = "image/png"
        Case ".gif"
            sID = "image/gif"
        Case ".jpg", "jpeg", ".jpe", "jfif"
            sID = "image/jpeg"
        Case ".tif", "tiff"
            sID = "image/tiff"
        Case ".bmp", ".dib"
            sID = "image/bmp"
        Case ".emf"
           sID = "image/x-emf"
        Case ".wmf"
           sID = "image/x-wmf"
        Case Else
            Exit Function
    End Select
    
    Call GdipCreateFromHDC(hdc, Graphics)
    Call GdipSetSmoothingMode(Graphics, SmoothingModeAntiAlias)
    Call GdipSaveGraphics(Graphics, State)

    If GdipCreateBitmapFromHBITMAP(hBitmap, 0&, lBitmap) = 0 Then
    
        If GetEncoderClsid(sID, PicEncoder) = True Then
            SavePictureFromHDC = (GdipSaveImageToFile(lBitmap, StrPtr(sFileName), PicEncoder, ByVal 0) = 0)
        End If
        GdipDisposeImage lBitmap
    End If
End Function

Private Function GetEncoderClsid(ByVal StrMimeType As String, ByRef ClassID As GUID) As Boolean
    
    Dim num As Long
    Dim Size As Long
    Dim imgCodecInfo() As ImageCodecInfo
    Dim lval As Long
    Dim buffer() As Byte

    GdipGetImageEncodersSize num, Size
    If Size Then
        ReDim imgCodecInfo(num) As ImageCodecInfo
        ReDim buffer(Size) As Byte

        GdipGetImageEncoders num, Size, buffer(0)
        CopyMemory imgCodecInfo(0), buffer(0), (Len(imgCodecInfo(0)) * num)

        For lval = 0 To num - 1
            If StrComp(GetStrFromPtrW(imgCodecInfo(lval).MimeType), StrMimeType, vbTextCompare) = 0 Then
                ClassID = imgCodecInfo(lval).ClassID
                GetEncoderClsid = True
                Exit For
            End If
        Next
        Erase imgCodecInfo
        Erase buffer
    End If
    
End Function

Private Function GetStrFromPtrW(ByVal lpszW As Long) As String
    
    Dim sRV As String

    sRV = String$(lStrlenW(ByVal lpszW) * 2, vbNullChar)
    WideCharToMultiByte CP_ACP, 0, ByVal lpszW, -1, ByVal sRV, Len(sRV), 0, 0
    GetStrFromPtrW = Left$(sRV, lStrlenW(StrPtr(sRV)))
    
End Function

Private Function StartUpGDIPlus() As Boolean
    
    Dim GdipStartupInput As GdiplusStartupInput
    
    GdipStartupInput.GdiplusVersion = 1
    GdipInitialized = (GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0&) = 0)
    StartUpGDIPlus = GdipInitialized
End Function

Private Sub ShutdownGDIPlus()
    
    If GdipInitialized Then
        GdiplusShutdown GdipToken
    End If
    
End Sub


