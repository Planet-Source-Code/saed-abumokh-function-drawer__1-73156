VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum CharSets
    ANSICharSet = 0
    DefaultCharSet = 1
    SymbolCharSet = 2
    MacCharSet = 77
    ShiftjisCharSet = 128
    HangeulCharSet = 129
    HangulCharSet = 129
    JohabCharSet = 130
    ChineseGB2312CharSet = 134
    ChineseBig5CharSet = 136
    GreekCharSet = 161
    TurkishCharSet = 162
    VietnameseCharSet = 163
    HebrewCharSet = 177
    ArabicCharSet = 178
    BalticCharSet = 186
    RussianCharSet = 204
    ThaiCharSet = 222
    EastEuropeCharSet = 238
    OEMCharSet = 255
End Enum

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

Private Type ChooseFont
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hdc As Long                '  printer DC/IC or NULL
        lpLogFont As Long
        iPointSize As Long         '  10 * Size in points of selected font
        Flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESize or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt Size allowed &
        nSizeMax As Long           '  max pt Size allowed if
                                       '    CF_LIMITSize is used
End Type

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long

Const CF_SCREENFONTS = &H1
Const CF_PRINTERFONTS = &H2
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_SHOWHELP = &H4&
Const CF_ENABLEHOOK = &H8&
Const CF_ENABLETEMPLATE = &H10&
Const CF_ENABLETEMPLATEHANDLE = &H20&
Const CF_INITTOLOGFONTStrUCT = &H40&
Const CF_USESTYLE = &H80&
Const CF_EFFECTS = &H100&
Const CF_APPLY = &H200&
Const CF_ANSIONLY = &H400&
Const CF_NOVECTORFONTS = &H800&
Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Const CF_NOSIMULATIONS = &H1000&
Const CF_LIMITSize = &H2000&
Const CF_FIXEDPITCHONLY = &H4000&
Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Const CF_FORCEFONTEXIST = &H10000
Const CF_SCALABLEONLY = &H20000
Const CF_TTONLY = &H40000
Const CF_NOFACESEL = &H80000
Const CF_NOSTYLESEL = &H100000
Const CF_NOSizeSEL = &H200000
Const CF_SELECTSCRIPT = &H400000
Const CF_NOSCRIPTSEL = &H800000
Const CF_NOVERTFONTS = &H1000000

Const SIMULATED_FONTTYPE = &H8000
Const PRINTER_FONTTYPE = &H4000
Const SCREEN_FONTTYPE = &H2000
Const BOLD_FONTTYPE = &H100
Const ITALIC_FONTTYPE = &H200
Const REGULAR_FONTTYPE = &H400

Dim cfOK As Boolean
Dim cfCancel As Boolean
Dim cfChooseFont As ChooseFont
Dim cfLogFont As LOGFONT
Dim cfhWndOwner As Long
Dim cfFontSize As Long
Dim cfFontColor As Long
Dim cfScreenFonts As Boolean
Dim cfPrinterFonts As Boolean
'Dim cfShowCommonPrinterAndScreenFonts As Boolean
Dim cfMinimumSize As Long
Dim cfMaximumSize As Long
Dim cfTrueTypeFontsOnly As Boolean
Dim cfShowEffects As Boolean
Dim cfANSICharSetOnly As Boolean
Dim cfFixedPitchFontsOnly As Boolean
Dim cfScalableFontsOnly As Boolean
Dim cfFontNotExistPrompt As Boolean
Dim cfShowHiddenFonts As Boolean
Dim cfInitializeFontProperties As Boolean
Dim cfLimitFontSize As Boolean
Dim cfDontShowVectorFonts As Boolean
Dim cfDontInitFontName As Boolean
Dim cfDontInitFontStyle As Boolean
Dim cfDontInitFontSize As Boolean
Dim cfNoVerticalFonts As Boolean
Dim cfShowHelpButton As Boolean
Dim cfSpecifyCharsets As Boolean
Dim cfDisableCharSet As Boolean

Friend Property Get hwndOwner() As Long
    hwndOwner = cfhWndOwner
End Property
    
Friend Property Let hwndOwner(ByVal vNewValue As Long)
    cfhWndOwner = vNewValue
End Property

Friend Property Get FontSize() As Long
    FontSize = Abs(cfLogFont.lfHeight * 3 / 4)
End Property

Friend Property Let FontSize(ByVal vNewValue As Long)
    cfLogFont.lfHeight = Abs(vNewValue * 4 / 3)
End Property

Friend Property Get FontColor() As Long
    FontColor = cfChooseFont.rgbColors
End Property

Friend Property Let FontColor(ByVal vNewValue As Long)
    cfChooseFont.rgbColors = vNewValue
End Property

Friend Property Get FontName() As String
    Dim i As Integer
    Dim retVal As String
    For i = 1 To 32
        retVal = retVal & Chr$(cfLogFont.lfFaceName(i))
    Next
    FontName = Replace(retVal, Chr$(0), "")
End Property

Friend Property Let FontName(ByVal vNewValue As String)
    Dim i As Integer
    For i = 1 To 32
        cfLogFont.lfFaceName(i) = 0
    Next
    For i = 1 To Len(vNewValue)
        cfLogFont.lfFaceName(i) = Asc(Mid(vNewValue, i, 1))
    Next
End Property

Friend Property Get FontBold() As Boolean
    If cfLogFont.lfWeight = 400 Then
        FontBold = False
    ElseIf cfLogFont.lfWeight = 700 Then
        FontBold = True
    End If
End Property

Friend Property Let FontBold(ByVal vNewValue As Boolean)
    If vNewValue = True Then
        cfLogFont.lfWeight = 700
    ElseIf vNewValue = False Then
        cfLogFont.lfWeight = 400
    End If
End Property

Friend Property Get FontItalic() As Boolean
    FontItalic = cfLogFont.lfItalic
End Property

Friend Property Let FontItalic(ByVal vNewValue As Boolean)
    cfLogFont.lfItalic = vNewValue
End Property

Friend Property Get FontUnderLine() As Boolean
    FontUnderLine = cfLogFont.lfUnderline
End Property

Friend Property Let FontUnderLine(ByVal vNewValue As Boolean)
    cfLogFont.lfUnderline = vNewValue
End Property

Friend Property Get FontStrikeThrough() As Boolean
    FontStrikeThrough = cfLogFont.lfStrikeOut
End Property

Friend Property Let FontStrikeThrough(ByVal vNewValue As Boolean)
    cfLogFont.lfStrikeOut = vNewValue
End Property

Friend Property Get FontCharSet() As Long
    FontCharSet = cfLogFont.lfCharSet
End Property

Friend Property Let FontCharSet(ByVal vNewValue As Long)
    cfLogFont.lfCharSet = vNewValue
End Property
Friend Property Get SpecifyCharsets() As Boolean
    SpecifyCharsets = cfSpecifyCharsets
End Property

Friend Property Let SpecifyCharsets(ByVal vNewValue As Boolean)
    cfSpecifyCharsets = vNewValue
End Property

Friend Property Get DisableCharSet() As Boolean
    DisableCharSet = cfDisableCharSet
End Property

Friend Property Let DisableCharSet(ByVal vNewValue As Boolean)
    cfDisableCharSet = vNewValue
End Property

Friend Property Get ShowScreenFonts() As Boolean
    ShowScreenFonts = cfScreenFonts
End Property

Friend Property Let ShowScreenFonts(ByVal vNewValue As Boolean)
    cfScreenFonts = vNewValue
End Property

Friend Property Get ShowPrinterFonts() As Boolean
    ShowPrinterFonts = cfPrinterFonts
End Property

Friend Property Let ShowPrinterFonts(ByVal vNewValue As Boolean)
    cfPrinterFonts = vNewValue
End Property

Friend Property Get MinimumSize() As Long
    MinimumSize = cfMinimumSize
End Property

Friend Property Let MinimumSize(ByVal vNewValue As Long)
    cfMinimumSize = vNewValue
End Property

Friend Property Get MaximumSize() As Long
    cfMaximumSize = MaximumSize
End Property

Friend Property Let MaximumSize(ByVal vNewValue As Long)
    cfMaximumSize = vNewValue
End Property

Friend Property Get TrueTypeFontsOnly() As Boolean
    TrueTypeFontsOnly = cfTrueTypeFontsOnly
End Property

Friend Property Let TrueTypeFontsOnly(ByVal vNewValue As Boolean)
    cfTrueTypeFontsOnly = vNewValue
End Property

Friend Property Get ANSICharSetOnly() As Boolean
    ANSICharSetOnly = cfANSICharSetOnly
End Property

Friend Property Let ANSICharSetOnly(ByVal vNewValue As Boolean)
    cfANSICharSetOnly = vNewValue
End Property
Friend Property Get ShowEffects() As Boolean
    ShowEffects = cfShowEffects
End Property

Friend Property Let ShowEffects(ByVal vNewValue As Boolean)
    cfShowEffects = vNewValue
End Property

Friend Property Get FixedPitchFontsOnly() As Boolean
    FixedPitchFontsOnly = cfFixedPitchFontsOnly
End Property

Friend Property Let FixedPitchFontsOnly(ByVal vNewValue As Boolean)
    cfFixedPitchFontsOnly = vNewValue
End Property

Friend Property Get ScalableFontsOnly() As Boolean
    ScalableFontsOnly = cfScalableFontsOnly
End Property

Friend Property Let ScalableFontsOnly(ByVal vNewValue As Boolean)
    cfScalableFontsOnly = vNewValue
End Property

Friend Property Get FontNotExistPrompt() As Boolean
    FontNotExistPrompt = cfFontNotExistPrompt
End Property

Friend Property Let FontNotExistPrompt(ByVal vNewValue As Boolean)
    cfFontNotExistPrompt = vNewValue
End Property

Friend Property Get ShowHiddenFonts() As Boolean
    ShowHiddenFonts = cfShowHiddenFonts
End Property

Friend Property Let ShowHiddenFonts(ByVal vNewValue As Boolean)
    cfShowHiddenFonts = vNewValue
End Property

Friend Property Get InitializeFontProperties() As Boolean
    InitializeFontProperties = cfInitializeFontProperties
End Property

Friend Property Let InitializeFontProperties(ByVal vNewValue As Boolean)
    cfInitializeFontProperties = vNewValue
End Property

Friend Property Get LimitFontSize() As Boolean
    LimitFontSize = cfLimitFontSize
End Property

Friend Property Let LimitFontSize(ByVal vNewValue As Boolean)
    cfLimitFontSize = vNewValue
End Property

Friend Property Get DontShowVectorFonts() As Boolean
    DontShowVectorFonts = cfDontShowVectorFonts
End Property

Friend Property Let DontShowVectorFonts(ByVal vNewValue As Boolean)
    cfDontShowVectorFonts = vNewValue
End Property

Friend Property Get NoVerticalFonts() As Boolean
    NoVerticalFonts = cfNoVerticalFonts
End Property

Friend Property Let NoVerticalFonts(ByVal vNewValue As Boolean)
    cfNoVerticalFonts = vNewValue
End Property

Friend Property Get DontInitFontName() As Boolean
    DontInitFontName = cfDontInitFontName
End Property

Friend Property Let DontInitFontName(ByVal vNewValue As Boolean)
    cfDontInitFontName = vNewValue
End Property

Friend Property Get DontInitFontStyle() As Boolean
    DontInitFontStyle = cfDontInitFontStyle
End Property

Friend Property Let DontInitFontStyle(ByVal vNewValue As Boolean)
    cfDontInitFontStyle = vNewValue
End Property

Friend Property Get DontInitFontSize() As Boolean
    DontInitFontSize = cfDontInitFontSize
End Property

Friend Property Let DontInitFontSize(ByVal vNewValue As Boolean)
    cfDontInitFontSize = vNewValue
End Property

Friend Property Get ShowHelpButton() As Boolean
    ShowHelpButton = cfShowHelpButton
End Property

Friend Property Let ShowHelpButton(ByVal vNewValue As Boolean)
    cfShowHelpButton = vNewValue
End Property

Friend Property Get OK() As Boolean
    OK = cfOK
End Property

Friend Property Get Cancel() As Boolean
    Cancel = cfCancel
End Property

Friend Property Get DialogTitle() As String
    DialogTitle = CommonDialogsHooks.CFDialogTitle
End Property

Friend Property Let DialogTitle(ByVal vNewValue As String)
   CommonDialogsHooks.CFDialogTitle = vNewValue
End Property

Friend Property Get OKButtonCaption() As String
    OKButtonCaption = CommonDialogsHooks.CFOKButtonCaption
End Property

Friend Property Let OKButtonCaption(ByVal vNewValue As String)
   CommonDialogsHooks.CFOKButtonCaption = vNewValue
End Property

Friend Property Get CancelButtonCaption() As String
    CancelButtonCaption = CommonDialogsHooks.CFCancelButtonCaption
End Property

Friend Property Let CancelButtonCaption(ByVal vNewValue As String)
   CommonDialogsHooks.CFCancelButtonCaption = vNewValue
End Property

Friend Property Get EnableColorComboBox() As Boolean
    EnableColorComboBox = CommonDialogsHooks.CFEnableColorComboBox
End Property

Friend Property Let EnableColorComboBox(ByVal vNewValue As Boolean)
   CommonDialogsHooks.CFEnableColorComboBox = vNewValue
End Property

Public Sub ShowDialog()
    Dim SaveLogFont As LOGFONT
    SaveLogFont = cfLogFont
    With cfChooseFont
        
        .Flags = CLng(Abs(cfScreenFonts * CF_SCREENFONTS) Or _
                Abs(cfPrinterFonts * CF_PRINTERFONTS) Or _
                Abs(cfShowHelpButton * CF_SHOWHELP) Or _
                Abs(cfTrueTypeFontsOnly * CF_TTONLY) Or _
                Abs(cfInitializeFontProperties * CF_INITTOLOGFONTStrUCT) Or _
                Abs(cfShowEffects * CF_EFFECTS) Or _
                Abs(cfANSICharSetOnly * CF_ANSIONLY) Or _
                Abs(cfDontShowVectorFonts * CF_NOVECTORFONTS) Or _
                Abs(cfLimitFontSize * CF_LIMITSize) Or _
                Abs(cfFixedPitchFontsOnly * CF_FIXEDPITCHONLY) Or _
                Abs(cfFontNotExistPrompt * CF_FORCEFONTEXIST) Or _
                Abs(cfScalableFontsOnly * CF_SCALABLEONLY) Or _
                Abs(cfDontInitFontName * CF_NOFACESEL) Or _
                Abs(cfDontInitFontStyle * CF_NOSTYLESEL) Or _
                Abs(cfDontInitFontSize * CF_NOSizeSEL) Or _
                Abs(cfDisableCharSet * CF_NOSCRIPTSEL) Or _
                Abs(cfSpecifyCharsets * CF_SELECTSCRIPT) Or _
                Abs(cfNoVerticalFonts * CF_NOVERTFONTS)) Or _
                CF_ENABLEHOOK
        .hwndOwner = cfhWndOwner
        .lpfnHook = GetProc(AddressOf CommonDialogsHooks.ChooseFontProc)
        .lpLogFont = VarPtr(cfLogFont)
        .lStructSize = Len(cfChooseFont)
        .nSizeMax = cfMaximumSize
        .nSizeMin = cfMinimumSize
        .rgbColors = cfFontColor
    End With
    
    cfOK = CBool(ChooseFont(cfChooseFont))
    cfCancel = Not (cfOK)
    cfFontColor = cfChooseFont.rgbColors
    
    If cfCancel Then cfLogFont = SaveLogFont
End Sub

Private Sub Class_Initialize()
    CommonDialogsHooks.CFDialogTitle = "Font"
    CommonDialogsHooks.CFOKButtonCaption = "OK"
    CommonDialogsHooks.CFCancelButtonCaption = "Cancel"
    CommonDialogsHooks.CFEnableColorComboBox = True
End Sub
