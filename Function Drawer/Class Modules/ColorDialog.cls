VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Const CC_RGBINIT = &H1
Const CC_FULLOPEN = &H2
Const CC_PREVENTFULLOPEN = &H4
Const CC_SHOWHELP = &H8
Const CC_ENABLEHOOK = &H10
Const CC_ENABLETEMPLATE = &H20
Const CC_ENABLETEMPLATEHANDLE = &H40
Const CC_SOLIDCOLOR = &H80
Const CC_ANYCOLOR = &H100

Private Type CHOOSECOLOR
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Dim ccRgbInit As Boolean
Dim ccFullOpen As Boolean
Dim ccPreventFullOpen As Boolean
Dim ccShowHelp As Boolean
Dim ccSolidColor As Boolean
Dim ccAnyColor As Boolean

'Dim ccCCDialogTitle As String
Dim cchWndOwner As Long
Dim ccColor As Long
Dim ccCustomColors() As Long

Dim ccOK As Boolean
Dim ccCancel As Boolean

Dim ccCHOOSECOLOR As CHOOSECOLOR

Friend Property Get hwndOwner() As Long
    hwndOwner = cchWndOwner
End Property

Friend Property Let hwndOwner(ByVal vNewValue As Long)
    cchWndOwner = vNewValue
End Property

Friend Property Get Color() As Long
    Color = ccColor
End Property

Friend Property Let Color(ByVal vNewValue As Long)
    ccColor = vNewValue
End Property

Public Sub SetCustomColors(ByRef CustomColors() As Long)
    ccCustomColors = CustomColors
End Sub

Public Function CustomColors() As Long()
    Dim retVal() As Long
    ReDim Preserve retVal(LBound(ccCustomColors) To 15 + LBound(ccCustomColors)) As Long
    retVal = ccCustomColors
    CustomColors = retVal
End Function

Friend Property Get InitializeRGB() As Boolean
    InitializeRGB = ccRgbInit
End Property

Friend Property Let InitializeRGB(ByVal vNewValue As Boolean)
    ccRgbInit = vNewValue
End Property

Friend Property Get FullOpen() As Boolean
     FullOpen = ccFullOpen
End Property

Friend Property Let FullOpen(ByVal vNewValue As Boolean)
    ccFullOpen = vNewValue
End Property

Friend Property Get PreventFullOpen() As Boolean
    PreventFullOpen = ccPreventFullOpen
End Property

Friend Property Let PreventFullOpen(ByVal vNewValue As Boolean)
    ccPreventFullOpen = vNewValue
End Property

Friend Property Get ShowHelpButton() As Boolean
    ShowHelpButton = ccShowHelp
End Property

Friend Property Let ShowHelpButton(ByVal vNewValue As Boolean)
    ccShowHelp = vNewValue
End Property

Friend Property Get SolidColor() As Boolean
     SolidColor = ccSolidColor
End Property

Friend Property Let SolidColor(ByVal vNewValue As Boolean)
    ccSolidColor = vNewValue
End Property

Friend Property Get AnyColor() As Boolean
    AnyColor = ccAnyColor
End Property

Friend Property Let AnyColor(ByVal vNewValue As Boolean)
    ccAnyColor = vNewValue
End Property

Public Sub ShowDialog()
    ReDim Preserve ccCustomColors(LBound(ccCustomColors) To 15 + LBound(ccCustomColors)) As Long
    With ccCHOOSECOLOR
        .Flags = Abs(ccRgbInit * CC_RGBINIT) Or _
                 Abs(ccFullOpen * CC_FULLOPEN) Or _
                 Abs(ccPreventFullOpen * CC_PREVENTFULLOPEN) Or _
                 Abs(ccShowHelp * CC_SHOWHELP) Or _
                 Abs(ccShowHelp * CC_SHOWHELP) Or _
                 Abs(ccSolidColor * CC_SOLIDCOLOR) Or _
                 Abs(ccAnyColor * CC_ANYCOLOR) Or _
                 CC_ENABLEHOOK
                
        .hInstance = 0
        .hwndOwner = cchWndOwner
        .lpCustColors = VarPtr(ccCustomColors(LBound(ccCustomColors)))
        .lStructSize = Len(ccCHOOSECOLOR)
        .rgbResult = ccColor
        .lpfnHook = GetProc(AddressOf ChooseColorProc)
    
        ccOK = CBool(CHOOSECOLOR(ccCHOOSECOLOR))
        ccCancel = Not (ccOK)
        
        ccColor = .rgbResult
    
    End With
End Sub

Friend Property Get OK() As Boolean
    OK = ccOK
End Property

Friend Property Get Cancel() As Boolean
    Cancel = ccCancel
End Property

Friend Property Get CCDialogTitle() As String
    CCDialogTitle = CommonDialogsHooks.CCDialogTitle
End Property

Friend Property Let CCDialogTitle(ByVal vNewValue As String)
    CommonDialogsHooks.CCDialogTitle = vNewValue
End Property

Private Sub Class_Initialize()
    CommonDialogsHooks.CCDialogTitle = "Color"
    CommonDialogsHooks.CCOKButtonCaption = "OK"
    CommonDialogsHooks.CCCancelButtonCaption = "Cancel"
    CommonDialogsHooks.CCDefineCustomColorsButtonCaption = "&Define Custom Colors >>"
    CommonDialogsHooks.CCAddToCustomColorsButtonCaption = "&Add to Custom Colors"
End Sub

Friend Property Get CCOKButtonCaption() As String
     CCOKButtonCaption = CommonDialogsHooks.CCOKButtonCaption
End Property

Friend Property Let CCOKButtonCaption(ByVal vNewValue As String)
     CommonDialogsHooks.CCOKButtonCaption = vNewValue
End Property

Friend Property Get CCCancelButtonCaption() As String
    CCCancelButtonCaption = CommonDialogsHooks.CCCancelButtonCaption
End Property

Friend Property Let CCCancelButtonCaption(ByVal vNewValue As String)
    CommonDialogsHooks.CCCancelButtonCaption = vNewValue
End Property

Friend Property Get CCDefineCustomColorsButtonCaption() As String
    CCDefineCustomColorsButtonCaption = CommonDialogsHooks.CCDefineCustomColorsButtonCaption
End Property

Friend Property Let CCDefineCustomColorsButtonCaption(ByVal vNewValue As String)
    CommonDialogsHooks.CCDefineCustomColorsButtonCaption = vNewValue
End Property

Friend Property Get CCAddToCustomColorsButtonCaption() As String
    CCDefineCustomColorsButtonCaption = CommonDialogsHooks.CCDefineCustomColorsButtonCaption
End Property

Friend Property Let CCAddToCustomColorsButtonCaption(ByVal vNewValue As String)
    CommonDialogsHooks.CCDefineCustomColorsButtonCaption = vNewValue
End Property
