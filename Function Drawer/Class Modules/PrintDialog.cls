VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PD_ALLPAGES = &H0
Private Const PD_SELECTION = &H1
Private Const PD_PAGENUMS = &H2
Private Const PD_NOSELECTION = &H4
Private Const PD_NOPAGENUMS = &H8
Private Const PD_COLLATE = &H10
Private Const PD_PRINTTOFILE = &H20
Private Const PD_PRINTSETUP = &H40
Private Const PD_NOWARNING = &H80
Private Const PD_RETURNDC = &H100
Private Const PD_RETURNIC = &H200
Private Const PD_RETURNDEFAULT = &H400
Private Const PD_SHOWHELP = &H800
Private Const PD_ENABLEPRINTHOOK = &H1000
Private Const PD_ENABLESETUPHOOK = &H2000
Private Const PD_ENABLEPRINTTEMPLATE = &H4000
Private Const PD_ENABLESETUPTEMPLATE = &H8000
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Private Const PD_USEDEVMODECOPIES = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Private Const PD_DISABLEPRINTTOFILE = &H80000
Private Const PD_HIDEPRINTTOFILE = &H100000
Private Const PD_NONETWORKBUTTON = &H200000

'  size of a device name string
Private Const CCHDEVICENAME = 32

'  size of a form name string
Private Const CCHFORMNAME = 32

' current version of specification
Private Const DM_SPECVERSION = &H320

' field selection bits
Private Const DM_ORIENTATION = &H1&
Private Const DM_PAPERSIZE = &H2&
Private Const DM_PAPERLENGTH = &H4&
Private Const DM_PAPERWIDTH = &H8&
Private Const DM_SCALE = &H10&
Private Const DM_COPIES = &H100&
Private Const DM_DEFAULTSOURCE = &H200&
Private Const DM_PRINTQUALITY = &H400&
Private Const DM_COLOR = &H800&
Private Const DM_DUPLEX = &H1000&
Private Const DM_YRESOLUTION = &H2000&
Private Const DM_TTOPTION = &H4000&
Private Const DM_COLLATE As Long = &H8000
Private Const DM_FORMNAME As Long = &H10000
Private Const DM_ALLFIELDS As Long = &H1FF1F

Private Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer 'XResolution
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long 'display devices only
        dmDisplayFrequency As Long 'display devices only
End Type

Private Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
        extra As String * 100
End Type

Private Type PrintDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        Flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

Dim pdPrintDlg As PrintDlg
Dim dmDevMode As DEVMODE
Dim dnDevNames As DEVNAMES

Public Enum Oreintations
    Portrait = 1
    LandScape = 2
End Enum

Public Enum PaperSizes
    PaperLetter = 1
    PaperLettersmall = 2
    PaperTabloid = 3
    PaperLedger = 4
    PaperLegal = 5
    PaperStatement = 6
    PaperExecutive = 7
    PaperA3 = 8
    PaperA4 = 9
    PaperA4small = 10
    PaperA5 = 11
    PaperB4 = 12
    PaperB5 = 13
    PaperFolio = 14
    PaperQuarto = 15
    Paper10x14 = 16
    Paper11x17 = 17
    PaperNote = 18
    PaperEnv9 = 19
    PaperEnv10 = 20
    PaperEnv11 = 21
    PaperEnv12 = 22
    PaperEnv14 = 23
    PaperCsheet = 24
    PaperDsheet = 25
    PaperEsheet = 26
    PaperEnvDl = 27
    PaperEnvC5 = 28
    PaperEnvC3 = 29
    PaperEnvC4 = 30
    PaperEnvC6 = 31
    PaperEnvC65 = 32
    PaperEnvB4 = 33
    PaperEnvB5 = 34
    PaperEnvB6 = 35
    PaperEnvItaly = 36
    PaperEnvMonarch = 37
    PaperEnvPersonal = 38
    PaperFanfoldUs = 39
    PaperFanfoldStdGerman = 40
    PaperFanfoldLglGerman = 41
    User = 256
End Enum

Public Enum PaperSources
    Onlyone = 1
    Lower = 2
    Middle = 3
    Manual = 4
    Envelope = 5
    Envmanual = 6
    Auto = 7
    Tractor = 8
    Smallfmt = 9
    Largefmt = 10
    Largecapacity = 11
    Cassette = 14
    User = 256
End Enum

Public Enum PrinterQualies
    Draft = (-1)
    Low = (-2)
    Medium = (-3)
    High = (-4)
End Enum

Public Enum ColorModes
    Monochrome = 1
    Color = 2
End Enum

Public Enum DoubleSidedPrintings
    Simplex = 1
    Vertical = 2
    horizontal = 3
End Enum

Public Enum TTFontsPrintingModes
    PrintAsGraphics = 1
    DownloadAsSoftFonts = 2
    substituteDeviceFonts = 3
End Enum

Dim dmDeviceName As String
Dim dmOrientation As Oreintations
Dim dmPaperSize As PaperSizes
Dim dmPaperWidth As Integer
Dim dmPaperHeight As Integer
Dim dmScalePercent As Integer
Dim dmCopies As Integer
Dim dmPaperSource As PaperSources
Dim dmPrintQuality As PrinterQualies
Dim dmXResolution As Integer
Dim dmYResolution As Integer
Dim dmColorMode As ColorModes
Dim dmDoubleSidedPrinting As DoubleSidedPrintings
Dim dmTTFontsPrintingMode As TTFontsPrintingModes
Dim dmCollate As Boolean
Dim dmPaperTypeName As String * CCHFORMNAME

Dim pdAllPagesSelected As Boolean
Dim pdSelectionSelected As Boolean
Dim pdPagesSelected As Boolean
Dim pdSelectionDisabled As Boolean
Dim pdPagesDisabled As Boolean
Dim pdCollateChecked As Boolean
Dim pdPrintToFileChecked As Boolean
Dim pdNoWarningNoDefaultPrinter As Boolean
Dim pdShowHelpButton As Boolean
Dim pdPrintToFileDisabled As Boolean
Dim pdPrintToFileHidden As Boolean
Dim pdNoNetworkButton As Boolean

Dim pdhWndOwner As Long
Dim pdhPrinterDC As Long
Dim pdFromPage As Integer
Dim pdToPage As Integer
Dim pdMinAllowedPages As Integer
Dim pdMaxAllowedPages As Integer
Dim pdOK As Boolean
Dim pdCancel As Boolean

Friend Property Get DeviceName() As String
    DeviceName = dmDeviceName
End Property

Friend Property Let DeviceName(ByVal vNewValue As String)
   dmDeviceName = vNewValue
End Property

Friend Property Get Orientation() As Oreintations
    Orientation = dmOrientation
End Property

Friend Property Let Orientation(ByVal vNewValue As Oreintations)
   dmOrientation = vNewValue
End Property

Friend Property Get PaperSize() As PaperSizes
    PaperSize = dmPaperSize
End Property

Friend Property Let PaperSize(ByVal vNewValue As PaperSizes)
   dmPaperSize = vNewValue
End Property

Friend Property Get PaperWidth() As Integer
    PaperWidth = dmPaperWidth
End Property

Friend Property Let PaperWidth(ByVal vNewValue As Integer)
   dmPaperWidth = vNewValue
End Property

Friend Property Get PaperHeight() As Integer
    PaperHeight = dmPaperHeight
End Property

Friend Property Let PaperHeight(ByVal vNewValue As Integer)
   dmPaperHeight = vNewValue
End Property

Friend Property Get ScalePercent() As Integer
    ScalePercent = dmScalePercent
End Property

Friend Property Let ScalePercent(ByVal vNewValue As Integer)
   dmScalePercent = vNewValue
End Property

Friend Property Get Copies() As Integer
    Copies = dmCopies
End Property

Friend Property Let Copies(ByVal vNewValue As Integer)
   dmCopies = vNewValue
End Property

Friend Property Get PaperSource() As PaperSources
    PaperSource = dmPaperSource
End Property

Friend Property Let PaperSource(ByVal vNewValue As PaperSources)
   dmPaperSource = vNewValue
End Property

Friend Property Get PrintQuality() As PrinterQualies
    PrintQuality = dmPrintQuality
End Property

Friend Property Let PrintQuality(ByVal vNewValue As PrinterQualies)
   dmPrintQuality = vNewValue
End Property

Friend Property Get XResolution() As Integer
    XResolution = dmXResolution
End Property

Friend Property Let XResolution(ByVal vNewValue As Integer)
   dmXResolution = vNewValue
End Property

Friend Property Get YResolution() As Integer
    YResolution = dmYResolution
End Property

Friend Property Let YResolution(ByVal vNewValue As Integer)
   dmYResolution = vNewValue
End Property

Friend Property Get ColorMode() As ColorModes
    ColorMode = dmColorMode
End Property

Friend Property Let ColorMode(ByVal vNewValue As ColorModes)
   dmColorMode = vNewValue
End Property

Friend Property Get DoubleSidedPrinting() As DoubleSidedPrintings
    DoubleSidedPrinting = dmDoubleSidedPrinting
End Property

Friend Property Let DoubleSidedPrinting(ByVal vNewValue As DoubleSidedPrintings)
   dmDoubleSidedPrinting = vNewValue
End Property

Friend Property Get TTFontsPrintingMode() As TTFontsPrintingModes
    TTFontsPrintingMode = dmTTFontsPrintingMode
End Property

Friend Property Let TTFontsPrintingMode(ByVal vNewValue As TTFontsPrintingModes)
   dmTTFontsPrintingMode = vNewValue
End Property

Friend Property Get Collate() As Boolean
    Collate = dmCollate
End Property

Friend Property Let Collate(ByVal vNewValue As Boolean)
   dmCollate = vNewValue
End Property

Friend Property Get PaperTypeName() As String
    PaperTypeName = Trim(Replace(dmPaperTypeName, Chr(0), ""))
End Property

Friend Property Let PaperTypeName(ByVal vNewValue As String)
   dmPaperTypeName = vNewValue
End Property

Friend Property Get hwndOwner() As Long
    hwndOwner = pdhWndOwner
End Property

Friend Property Let hwndOwner(ByVal vNewValue As Long)
   pdhWndOwner = vNewValue
End Property

Friend Property Get hPrinterDC() As Long
    hPrinterDC = pdhPrinterDC
End Property

Friend Property Let hPrinterDC(ByVal vNewValue As Long)
   pdhPrinterDC = vNewValue
End Property

Friend Property Get FromPage() As Integer
    FromPage = pdFromPage
End Property

Friend Property Let FromPage(ByVal vNewValue As Integer)
   pdFromPage = vNewValue
End Property

Friend Property Get ToPage() As Integer
    ToPage = pdToPage
End Property

Friend Property Let ToPage(ByVal vNewValue As Integer)
   pdToPage = vNewValue
End Property

Friend Property Get MinAllowedPages() As Integer
    MinAllowedPages = pdMinAllowedPages
End Property

Friend Property Let MinAllowedPages(ByVal vNewValue As Integer)
   pdMinAllowedPages = vNewValue
End Property

Friend Property Get MaxAllowedPages() As Integer
    MaxAllowedPages = pdMaxAllowedPages
End Property

Friend Property Let MaxAllowedPages(ByVal vNewValue As Integer)
   pdMaxAllowedPages = vNewValue
End Property

Friend Property Get AllPagesSelected() As Boolean
    AllPagesSelected = pdAllPagesSelected
End Property

Friend Property Let AllPagesSelected(ByVal vNewValue As Boolean)
   pdAllPagesSelected = vNewValue
End Property

Friend Property Get SelectionSelected() As Boolean
    SelectionSelected = pdSelectionSelected
End Property

Friend Property Let SelectionSelected(ByVal vNewValue As Boolean)
   pdSelectionSelected = vNewValue
End Property

Friend Property Get PagesSelected() As Boolean
    PagesSelected = pdPagesSelected
End Property

Friend Property Let PagesSelected(ByVal vNewValue As Boolean)
   pdPagesSelected = vNewValue
End Property

Friend Property Get SelectionDisabled() As Boolean
    SelectionDisabled = pdSelectionDisabled
End Property

Friend Property Let SelectionDisabled(ByVal vNewValue As Boolean)
   pdSelectionDisabled = vNewValue
End Property

Friend Property Get PagesDisabled() As Boolean
    PagesDisabled = pdPagesDisabled
End Property

Friend Property Let PagesDisabled(ByVal vNewValue As Boolean)
   pdPagesDisabled = vNewValue
End Property

Friend Property Get PrintToFileChecked() As Boolean
    PrintToFileChecked = pdPrintToFileChecked
End Property

Friend Property Let PrintToFileChecked(ByVal vNewValue As Boolean)
   pdPrintToFileChecked = vNewValue
End Property

Friend Property Get CollateChecked() As Boolean
    CollateChecked = pdCollateChecked
End Property

Friend Property Let CollateChecked(ByVal vNewValue As Boolean)
   pdCollateChecked = vNewValue
End Property

Friend Property Get NoWarningNoDefaultPrinter() As Boolean
    NoWarningNoDefaultPrinter = pdNoWarningNoDefaultPrinter
End Property

Friend Property Let NoWarningNoDefaultPrinter(ByVal vNewValue As Boolean)
   pdNoWarningNoDefaultPrinter = vNewValue
End Property

Friend Property Get ShowHelpButton() As Boolean
    ShowHelpButton = pdShowHelpButton
End Property

Friend Property Let ShowHelpButton(ByVal vNewValue As Boolean)
   pdShowHelpButton = vNewValue
End Property

Friend Property Get PrintToFileDisabled() As Boolean
    PrintToFileDisabled = pdPrintToFileDisabled
End Property

Friend Property Let PrintToFileDisabled(ByVal vNewValue As Boolean)
   pdPrintToFileDisabled = vNewValue
End Property

Friend Property Get PrintToFileHidden() As Boolean
    PrintToFileHidden = pdPrintToFileHidden
End Property

Friend Property Let PrintToFileHidden(ByVal vNewValue As Boolean)
   pdPrintToFileHidden = vNewValue
End Property

Friend Property Get NoNetworkButton() As Boolean
    NoNetworkButton = pdNoNetworkButton
End Property

Friend Property Let NoNetworkButton(ByVal vNewValue As Boolean)
   pdNoNetworkButton = vNewValue
End Property

Friend Property Get OK() As Boolean
    OK = pdOK
End Property

Friend Property Get Cancel() As Boolean
    Cancel = pdCancel
End Property

Public Sub ShowDialog()
    pdOK = ShowPrinter(False)
    pdCancel = Not (pdOK)
    SetUserSelections
End Sub

Public Sub ShowPrintSetupDialog()
    pdOK = ShowPrinter(True)
    pdCancel = Not (pdOK)
    SetUserSelections
End Sub

Private Function ShowPrinter(ByVal ShowPrinterSetup As Boolean) As Boolean
    Dim hDevMode As Long, hDevNames As Long
    With dmDevMode
        .dmDeviceName = dmDeviceName
        .dmOrientation = dmOrientation
        .dmPaperSize = dmPaperSize
        .dmPaperWidth = dmPaperWidth
        .dmPaperLength = dmPaperHeight
        .dmScale = dmScalePercent
        .dmCopies = dmCopies
        .dmDefaultSource = dmPaperSource
        .dmPrintQuality = dmPrintQuality
        If dmXResolution > 0 Then .dmPrintQuality = dmXResolution
        .dmYResolution = dmYResolution
        .dmColor = dmColorMode
        .dmDuplex = dmDoubleSidedPrinting
        .dmTTOption = dmTTFontsPrintingMode
        .dmCollate = dmCollate
        .dmFormName = dmPaperTypeName
        .dmSize = Len(dmDevMode)
        
        
        '.dmFields = DM_ORIENTATION Or DM_DUPLEX
        .dmFields = DM_COLLATE Or DM_COLOR Or DM_COPIES Or DM_DEFAULTSOURCE Or DM_DUPLEX Or DM_FORMNAME Or _
                    DM_FORMNAME Or DM_ORIENTATION Or DM_PAPERLENGTH Or DM_PAPERSIZE Or DM_PAPERWIDTH Or _
                    DM_PRINTQUALITY Or DM_SCALE Or DM_TTOPTION Or DM_YRESOLUTION
        
    End With
    
    With dnDevNames
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
        .extra = Printer.DriverName & Chr(0) & Printer.DeviceName & Chr(0) & Printer.Port & Chr(0)
    End With

    With pdPrintDlg
        .Flags = Abs(pdAllPagesSelected * PD_ALLPAGES) Or _
                 Abs(pdSelectionSelected * PD_SELECTION) Or _
                 Abs(pdPagesSelected * PD_PAGENUMS) Or _
                 Abs(pdSelectionDisabled * PD_NOSELECTION) Or _
                 Abs(pdPagesDisabled * PD_NOPAGENUMS) Or _
                 Abs(pdPrintToFileChecked * PD_PRINTTOFILE) Or _
                 Abs(pdCollateChecked * PD_COLLATE) Or _
                 Abs(pdNoWarningNoDefaultPrinter * PD_NOWARNING) Or _
                 Abs(pdShowHelpButton * PD_SHOWHELP) Or _
                 Abs(pdPrintToFileDisabled * PD_DISABLEPRINTTOFILE) Or _
                 Abs(pdPrintToFileHidden * PD_HIDEPRINTTOFILE) Or _
                 Abs(pdNoNetworkButton * PD_NONETWORKBUTTON) Or _
                 Abs(ShowPrinterSetup * PD_PRINTSETUP) Or _
                 PD_RETURNDC Or PD_USEDEVMODECOPIESANDCOLLATE
        .hdc = pdhPrinterDC
        .hDevMode = VarPtr(dmDevMode)
        .hDevMode = VarPtr(dnDevNames)
        .hwndOwner = pdhWndOwner
        .lStructSize = Len(pdPrintDlg)
        .nCopies = dmCopies
        .nFromPage = pdFromPage
        .nToPage = pdToPage
        .nMaxPage = pdMaxAllowedPages
        .nMinPage = pdMinAllowedPages
    End With

    pdPrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(dmDevMode))
    hDevMode = GlobalLock(pdPrintDlg.hDevMode)
    If hDevMode > 0 Then
        CopyMemory ByVal hDevMode, dmDevMode, Len(dmDevMode)
        GlobalUnlock pdPrintDlg.hDevMode
    End If
    
    pdPrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(dnDevNames))
    hDevNames = GlobalLock(pdPrintDlg.hDevNames)
    If hDevNames > 0 Then
        CopyMemory ByVal hDevNames, dnDevNames, Len(dnDevNames)
        GlobalUnlock hDevNames
    End If
    
    ShowPrinter = CBool(PrintDlg(pdPrintDlg))
    If ShowPrinter = True Then
        hDevNames = GlobalLock(pdPrintDlg.hDevNames)
        CopyMemory dnDevNames, ByVal hDevNames, 45
        GlobalUnlock hDevNames
        GlobalFree pdPrintDlg.hDevNames
        hDevMode = GlobalLock(pdPrintDlg.hDevMode)
        CopyMemory dmDevMode, ByVal hDevMode, Len(dmDevMode)
        GlobalUnlock pdPrintDlg.hDevMode
        GlobalFree pdPrintDlg.hDevMode
        dmDeviceName = Left(dmDevMode.dmDeviceName, InStr(dmDevMode.dmDeviceName, Chr$(0)) - 1)
    End If

End Function

Private Sub Class_Initialize()
        dmDeviceName = Printer.DeviceName
        dmOrientation = Printer.Orientation
        dmPaperSize = Printer.PaperSize
        dmPaperWidth = Printer.Width
        dmPaperHeight = Printer.Height
        dmScalePercent = Printer.Zoom
        dmCopies = Printer.Copies
        dmPaperSource = Printer.PaperBin
        dmPrintQuality = Printer.PrintQuality
        dmColorMode = Printer.ColorMode
        dmDoubleSidedPrinting = Printer.Duplex
        dmCollate = 1
End Sub
Private Sub SetUserSelections()
    With dmDevMode
        dmOrientation = .dmOrientation
        dmPaperSize = .dmPaperSize
        dmPaperWidth = .dmPaperWidth
        dmPaperHeight = .dmPaperLength
        dmScalePercent = .dmScale
        dmCopies = .dmCopies
        dmPaperSource = .dmDefaultSource
        dmPrintQuality = .dmPrintQuality
        If dmPrintQuality < 0 Then dmXResolution = .dmPrintQuality
        dmYResolution = .dmYResolution
        dmColorMode = .dmColor
        dmDoubleSidedPrinting = .dmDuplex
        dmTTFontsPrintingMode = .dmTTOption
        dmCollate = .dmCollate
        dmPaperTypeName = .dmFormName
    End With
    With pdPrintDlg
        pdFromPage = .nFromPage
        pdMaxAllowedPages = .nMaxPage
        pdMinAllowedPages = .nMinPage
        pdToPage = .nToPage
        
        pdSelectionSelected = (.Flags And PD_SELECTION) = PD_SELECTION
        pdPagesSelected = (.Flags And PD_PAGENUMS) = PD_PAGENUMS
        pdAllPagesSelected = Not (pdSelectionSelected Or pdPagesSelected)
        pdSelectionDisabled = (.Flags And PD_NOSELECTION) = PD_NOSELECTION
        pdPagesDisabled = (.Flags And PD_NOPAGENUMS) = PD_NOPAGENUMS
        pdPrintToFileChecked = (.Flags And PD_PRINTTOFILE) = PD_PRINTTOFILE
        pdCollateChecked = (.Flags And PD_COLLATE) = PD_COLLATE
        pdNoWarningNoDefaultPrinter = (.Flags And PD_NOWARNING) = PD_NOWARNING
        pdShowHelpButton = (.Flags And PD_SHOWHELP) = PD_SHOWHELP
        pdPrintToFileDisabled = (.Flags And PD_DISABLEPRINTTOFILE) = PD_DISABLEPRINTTOFILE
        pdPrintToFileHidden = (.Flags And PD_HIDEPRINTTOFILE) = PD_HIDEPRINTTOFILE
        pdNoNetworkButton = (.Flags And PD_NONETWORKBUTTON) = PD_NONETWORKBUTTON
    End With
End Sub


