VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Function Drawer"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   240
      Left            =   30
      TabIndex        =   6
      Top             =   7470
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   7425
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2910
            Text            =   "For Help, press F1.    "
            TextSave        =   "For Help, press F1.    "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   2910
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   2910
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   2910
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbFunction 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":08CA
      Left            =   540
      List            =   "frmMain.frx":08CC
      TabIndex        =   3
      Text            =   "X"
      Top             =   6960
      Width           =   6795
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   6960
      Width           =   855
   End
   Begin VB.PictureBox picFunction 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   11.88
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   14.367
      TabIndex        =   0
      Top             =   120
      Width           =   8145
   End
   Begin VB.Timer tmrFunctionDetailsOnMenu 
      Interval        =   250
      Left            =   3720
      Top             =   3480
   End
   Begin VB.PictureBox picMenuColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   990
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin ComctlLib.ImageList imgLst 
      Left            =   3960
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   2
      Top             =   7050
      Width           =   60
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFileSaveAsPicture 
         Caption         =   "Save As Picture"
      End
      Begin VB.Menu mnuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCopyGraphImage 
         Caption         =   "Copy Graph Image"
      End
      Begin VB.Menu mnuEditCopyFunctions 
         Caption         =   "Copy Functions"
      End
      Begin VB.Menu mnuEditPasteFunctions 
         Caption         =   "Paste Functions"
      End
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "Function"
      Begin VB.Menu mnuFunctionDrawTangent 
         Caption         =   "Draw Tangent"
      End
      Begin VB.Menu mnuFunctionAddNewConstant 
         Caption         =   "Add New Constant"
      End
      Begin VB.Menu mnuFunctionSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionSmoothDrawing 
         Caption         =   "Smooth Drawing"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Graph Options"
      Begin VB.Menu mnuOptionsChangeColor 
         Caption         =   "&Change Color"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuOptionsDrawStyle 
         Caption         =   "Draw Style"
         Begin VB.Menu mnuOptionsDrawStyles 
            Caption         =   "Dotted Soft"
            Index           =   1
         End
         Begin VB.Menu mnuOptionsDrawStyles 
            Caption         =   "Dotted More"
            Index           =   2
         End
         Begin VB.Menu mnuOptionsDrawStyles 
            Caption         =   "Dashed (or Check board)"
            Index           =   3
         End
         Begin VB.Menu mnuOptionsDrawStyles 
            Caption         =   "Normal"
            Index           =   4
         End
      End
      Begin VB.Menu mnuOptionsBorderWidth 
         Caption         =   "Border Width"
         Begin VB.Menu mnuOptionsBorderWidths 
            Caption         =   "1 pt"
            Index           =   1
         End
         Begin VB.Menu mnuOptionsBorderWidths 
            Caption         =   "2 pt"
            Index           =   2
         End
         Begin VB.Menu mnuOptionsBorderWidths 
            Caption         =   "3 pt"
            Index           =   3
         End
         Begin VB.Menu mnuOptionsBorderWidths 
            Caption         =   "4 pt"
            Index           =   4
         End
         Begin VB.Menu mnuOptionsBorderWidths 
            Caption         =   "5 pt"
            Index           =   5
         End
         Begin VB.Menu mnuOptionsBorderWidths 
            Caption         =   "7 pt"
            Index           =   6
         End
         Begin VB.Menu mnuOptionsBorderWidths 
            Caption         =   "10 pt"
            Index           =   7
         End
         Begin VB.Menu mnuOptionsBorderWidths 
            Caption         =   "15 pt"
            Index           =   8
         End
      End
      Begin VB.Menu mnuFileOptionsSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsGrid 
         Caption         =   "Grid And Axises Options..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Function Drawer"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   ""
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sc As New MSScriptControl.ScriptControl
Public ConstsAdded As String

Dim lastX As Single, lastY As Single
Dim Colors() As Long
Dim BorderWidths() As Single
Dim PenStyles() As Integer

Dim pd As New PrintDialog
Dim ColorDlg As New ColorDialog
Dim CustColors(1 To 16) As Long

Dim PenHatchStyle As PenHatchStyles

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Dim FileSaveCounter As Long
Public IsSaved As Boolean
Dim CurrentFileName As String
Dim CurrentFilePath As String
Dim PromptSave As Boolean

Dim DrawTangentMode As Boolean

Dim MenuIcons() As String

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Sub InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX)

Private Type tagINITCOMMONCONTROLSEX
  dwSize As Long
  dwICC As Long
End Type

Dim ms As New MenuStyle

Dim SaveAndUnload As Boolean

Private Sub cmbFunction_Click()
    TrimComboItems cmbFunction
End Sub

Private Sub cmdDraw_Click()
    Dim i As Integer
    
    TrimComboItems cmbFunction
    
    For i = 0 To cmbFunction.ListCount - 1
        If cmbFunction.list(i) = cmbFunction.Text Then GoTo DontAdd
    Next
    cmbFunction.AddItem cmbFunction.Text
DontAdd:
    picFunction.Picture = Nothing
    DrawFunctions
    
    On Error Resume Next
    ClearMnuOptionsDrawStyle
    ClearMnuOptionsBorderWidths
    If cmbFunction.ListIndex = -1 Then cmbFunction.ListIndex = cmbFunction.ListCount - 1
    mnuOptionsDrawStyles(HatchStyleToIndex(PenStyles(cmbFunction.ListIndex))).Checked = True
    mnuOptionsBorderWidths(BorderWidthToIndex(BorderWidths(cmbFunction.ListIndex))).Checked = True
    
    If BorderWidths(cmbFunction.ListIndex) = 0 Then _
    BorderWidths(cmbFunction.ListIndex) = IndexToBorderWidth(1)
    
    cmbFunction.SetFocus
    picFunction.Picture = picFunction.Image
End Sub

Public Sub DrawFunctions()
    Dim RGBColor As RGB
    Dim i As Integer
    On Error Resume Next
    picFunction.Cls
    DisplayAxisesAndGrid picFunction
    
    ReDim Preserve Colors(0 To cmbFunction.ListCount - 1)
    ReDim Preserve PenStyles(0 To cmbFunction.ListCount - 1)
    ReDim Preserve BorderWidths(0 To cmbFunction.ListCount - 1)
    
    ProgressBar.Max = cmbFunction.ListCount - 1
    ProgressBar.Visible = True
    For i = 0 To cmbFunction.ListCount - 1
        ProgressBar.Value = i
        RGBColor = ColorToRGB(Colors(i))
        DrawBorder picFunction
        Printer.KillDoc
        DrawFunction picFunction, ConstsAdded, cmbFunction.list(i), 0, 0, picFunction.ScaleWidth, picFunction.ScaleHeight, NewARGB(255, RGBColor.r, RGBColor.g, RGBColor.b), BorderWidths(i), PenStyles(i), mnuFunctionSmoothDrawing.Checked
        DrawBorder picFunction
    Next
    ProgressBar.Visible = False
    picFunction.Refresh
    Refresh
End Sub

Private Sub PrintPicture()
    Dim Inch As Single
    Printer.KillDoc
    Printer.ScaleMode = vbPixels
    Inch = Printer.ScaleX(1, vbInches, Printer.ScaleMode)
    Printer.PaintPicture picFunction.Image, ((Printer.Width / Printer.TwipsPerPixelX) - Printer.ScaleX(picFunction.Width, picFunction.ScaleMode, Printer.ScaleMode)) / 2, Inch
    Printer.EndDoc
    Me.Caption = Printer.Width / Printer.TwipsPerPixelX / Printer.TwipsPerPixelX
    Me.Caption = Printer.TwipsPerPixelX
    Me.Caption = Inch
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbShiftMask + vbCtrlMask And KeyCode = vbKeyS Then
        If mnuFileSaveAs.Enabled = True Then mnuFileSaveAs_Click
    ElseIf Shift = vbShiftMask + vbCtrlMask And KeyCode = vbKeyC Then
        If mnuEditCopyFunctions.Enabled = True Then mnuEditCopyFunctions_Click
    ElseIf Shift = vbShiftMask + vbCtrlMask And KeyCode = vbKeyV Then
        If mnuEditPasteFunctions.Enabled = True Then mnuEditPasteFunctions_Click
    ElseIf Shift = vbAltMask And KeyCode = vbKeyT Then
        If mnuFunctionDrawTangent.Enabled = True Then mnuFunctionDrawTangent_Click
    ElseIf KeyCode = vbKeyEscape Then
        If mnuOther.Visible = True Then mnuOther_Click
    End If
End Sub

Private Sub Form_Initialize()
  Dim icc As tagINITCOMMONCONTROLSEX
  
  icc.dwICC = &HFF
  icc.dwSize = 8
  
  Call InitCommonControlsEx(icc)
End Sub

Public Sub Menus_MouseSelect(ByVal Index As Long, ByVal HasSubMenus As Boolean, ByVal MenuCaption As String, ByVal HotKey As String)
    On Error Resume Next
    If HasSubMenus = False Then
        Select Case Split(MenuCaption, vbTab)(0)
        Case Is = mnuFileNew.Caption
            StatusBar.Panels(1).Text = "Creates a new math functions document."
        Case Is = mnuFileOpen.Caption
            StatusBar.Panels(1).Text = "Opens an existing document."
        Case Is = mnuFileSave.Caption
            StatusBar.Panels(1).Text = "Saves the current document's changes."
        Case Is = "Save As..."
            StatusBar.Panels(1).Text = "Saves the current document to file."
        Case Is = mnuFileSaveAsPicture.Caption
            StatusBar.Panels(1).Text = "Saves the written graphs as a picture."
        Case Is = mnuFilePrint.Caption
            StatusBar.Panels(1).Text = "Prints the written graphs."
        Case Is = "Close"
            StatusBar.Panels(1).Text = "Quits Function Drawer; prompts you to save changes on the documents."
        Case Is = mnuEditCopyGraphImage.Caption
            StatusBar.Panels(1).Text = "Copies the graph as a picture to the clipboard."
        Case Is = "Copy Functions"
            StatusBar.Panels(1).Text = "Copies the functions formulas to the clipboard."
        Case Is = "Paste Functions"
            StatusBar.Panels(1).Text = "Pastes the functions formulas from the clipboard."
        Case Is = "Draw Tangent"
            StatusBar.Panels(1).Text = "Draws a tangent to the current function."
        Case Is = mnuFunctionAddNewConstant.Caption
            StatusBar.Panels(1).Text = "Adds a new contant to the current constants."
        Case Is = mnuFunctionSmoothDrawing.Caption
            StatusBar.Panels(1).Text = "Toggles smooth curve drawing for graphs."
        Case Is = mnuOptionsChangeColor.Caption
            StatusBar.Panels(1).Text = "Specifies the current graph color."
        Case Is = mnuOptionsDrawStyles(1).Caption
            StatusBar.Panels(1).Text = "Makes the graph line is softly dotted."
        Case Is = mnuOptionsDrawStyles(2).Caption
            StatusBar.Panels(1).Text = "Makes the graph line is dotted more."
        Case Is = mnuOptionsDrawStyles(3).Caption
            StatusBar.Panels(1).Text = "Makes the graph line is dashed, or filled with a checkbord pattern."
        Case Is = mnuOptionsDrawStyles(4).Caption
            StatusBar.Panels(1).Text = "Removes the line style of the line graph."
        Case Is = mnuOptionsBorderWidths(1).Caption
            StatusBar.Panels(1).Text = "Sets the line graph width to 1 point"
        Case Is = mnuOptionsBorderWidths(2).Caption
            StatusBar.Panels(1).Text = "Sets the line graph width to 2 points"
        Case Is = mnuOptionsBorderWidths(3).Caption
            StatusBar.Panels(1).Text = "Sets the line graph width to 3 points"
        Case Is = mnuOptionsBorderWidths(4).Caption
            StatusBar.Panels(1).Text = "Sets the line graph width to 4 points"
        Case Is = mnuOptionsBorderWidths(5).Caption
            StatusBar.Panels(1).Text = "Sets the line graph width to 5 points"
        Case Is = mnuOptionsBorderWidths(6).Caption
            StatusBar.Panels(1).Text = "Sets the line graph width to 7 points"
        Case Is = mnuOptionsBorderWidths(7).Caption
            StatusBar.Panels(1).Text = "Sets the line graph width to 10 points"
        Case Is = mnuOptionsBorderWidths(8).Caption
            StatusBar.Panels(1).Text = "Sets the line graph width to 15 points"
        Case Is = mnuOptionsGrid.Caption
            StatusBar.Panels(1).Text = "Changes the axises, grid, and the numbers' graphical styles."
        Case Is = mnuHelpContents.Caption
            StatusBar.Panels(1).Text = "Views the help contents of Function Drawer."
        Case Is = mnuHelpAbout.Caption
            StatusBar.Panels(1).Text = "Displays Function Drawer's info, version number, and copyright."
        Case Else
            StatusBar.Panels(1).Text = "For Help, press F1.    "
        End Select
    Else
        If Index < 1 Then
            StatusBar.Panels(1).Text = "For Help, press F1.    "
        Else
            StatusBar.Panels(1).Text = ""
        End If
    End If

End Sub

Private Sub Form_Load()

    frmSplash.Show
    frmSplash.Refresh

    Load frmAbout
    Load frmAddNewConst
    Load frmGridAndAxises
    Load frmHelp


    SetAntiAliasedFontControls frmAbout
    SetAntiAliasedFontControls frmAddNewConst
    SetAntiAliasedFontControls frmGridAndAxises
    SetAntiAliasedFontControls frmHelp
    SetAntiAliasedFontControls frmSplash
    SetAntiAliasedFontControls Me
    DoEvents

    
    DisplayAxisesAndGrid picFunction
    sc.Language = "VBScript"
    sc.AddCode _
    LoadEncryptedFile(EncryptionCode, App.path & "\Data\More Functions.nmf")

    Label1.Caption = Chr$(131) & "(x)="

    cmbFunction.AddItem cmbFunction.Text
    cmbFunction.ListIndex = 0

    Dim i As Integer
    For i = 1 To 16
        CustColors(i) = QBColor(i - 1)
    Next
    ColorDlg.SetCustomColors CustColors

    InitAxsisedAndGrid

    CurrentFileName = "Functions 1.mfd"

    SetParent ProgressBar.hwnd, StatusBar.hwnd

    Me.ScaleMode = vbPixels
    ProgressBar.Left = 3
    ProgressBar.Top = 3
    Me.ScaleMode = vbCentimeters

    Dim g As New Gradient
    Dim dc As New DeviceContext
    dc.Create 24, 1, 26
    g.Rectangle dc.Handle, 0, 0, dc.Width, dc.Height, RGB(215, 233, 255), RGB(215, 233, 255), True
    SetMenuBarBackground Me.hwnd, dc.ConvertToBitmap(0, 0, dc.Width, dc.Height)
    dc.Dispose

    InitMenu
    mnuOther.Visible = False

    Printer.ScaleMode = vbPixels

    Unload frmSplash

End Sub

Private Sub InitAxsisedAndGrid()
    With MainAxisesStyle
        .BorderColor = 0
        .BorderWidth = 2
        .DrawStyle = DrawStyleConstants.vbSolid
        .Visible = True
    End With
    
    With MainGridStyle
        .BorderColor = RGB(128, 128, 128)
        .BorderWidth = 1
        .DrawStyle = DrawStyleConstants.vbDot
        .Visible = True
    End With
    
    With MainNumbersFont
        .Bold = True
        .Charset = CharSets.DefaultCharSet
        .Name = "Tahoma"
        .Size = 8
        .Weight = 700
    End With
    
    MainNumbersColor = vbBlue
    MainNumbersVisible = True
    
End Sub

Private Sub InitMenu()

On Error Resume Next
    ms.Clear
    
    mnuFileClose.Caption = mnuFileClose.Caption & vbTab & "Alt+F4"
    mnuFileSaveAs.Caption = mnuFileSaveAs.Caption & vbTab & "Ctrl+Shift+S"
    mnuEditCopyFunctions.Caption = mnuEditCopyFunctions.Caption & vbTab & "Ctrl+Shift+C"
    mnuEditPasteFunctions.Caption = mnuEditPasteFunctions.Caption & vbTab & "Ctrl+Shift+V"
    mnuFunctionDrawTangent.Caption = mnuFunctionDrawTangent.Caption & vbTab & "Alt+T"

    ms.hwnd = Me.hwnd
    ms.MenuFont = Me.Font
    
    ReDim MenuIcons(0 To ms.MenuCount + 1)
    Dim IconsFolder As String
    IconsFolder = Replace(App.path & "\", "\\", "\") & "data\pictures\"
    
    MenuIcons(2) = IconsFolder & "new.png"
    MenuIcons(4) = IconsFolder & "open.png"
    MenuIcons(5) = IconsFolder & "save.png"
    MenuIcons(6) = IconsFolder & "saveas.png"
    MenuIcons(7) = IconsFolder & "saveaspicture.png"
    MenuIcons(9) = IconsFolder & "print.png"
    MenuIcons(11) = IconsFolder & "close.png"
    MenuIcons(13) = IconsFolder & "copygraphimage.png"
    MenuIcons(14) = IconsFolder & "copyfunctions.png"
    MenuIcons(15) = IconsFolder & "pastefunctions.png"
    MenuIcons(17) = IconsFolder & "drawtangent.png"
    MenuIcons(18) = IconsFolder & "addnewconst.png"
    MenuIcons(22) = IconsFolder & "changecolor.png"
    MenuIcons(24) = IconsFolder & "dottedsoft.png"
    MenuIcons(25) = IconsFolder & "dottedmore.png"
    MenuIcons(26) = IconsFolder & "dash.png"
    MenuIcons(27) = IconsFolder & "normal.png"
    MenuIcons(29) = IconsFolder & "1pt.png"
    MenuIcons(30) = IconsFolder & "2pt.png"
    MenuIcons(31) = IconsFolder & "3pt.png"
    MenuIcons(32) = IconsFolder & "4pt.png"
    MenuIcons(33) = IconsFolder & "5pt.png"
    MenuIcons(34) = IconsFolder & "7pt.png"
    MenuIcons(35) = IconsFolder & "10pt.png"
    MenuIcons(36) = IconsFolder & "15pt.png"
    MenuIcons(38) = IconsFolder & "gridandaxisesoptions.png"
    MenuIcons(40) = IconsFolder & "help.png"
    MenuIcons(42) = IconsFolder & "about.png"
    
    ms.ImageFiles = MenuIcons
    ms.ImageWidth = 16
    ms.ImageHeight = 16
    ms.SetStyle
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    Me.ScaleMode = vbPixels
    
    picFunction.Width = Me.ScaleWidth - picFunction.Left - 10
    picFunction.Height = Me.ScaleHeight - picFunction.Top - StatusBar.Height - 10 - cmbFunction.Height - 10
    
    cmbFunction.Width = Me.ScaleWidth - cmbFunction.Left - 10 - cmdDraw.Width - 10
    cmbFunction.Top = Me.ScaleHeight - 10 - StatusBar.Height - cmbFunction.Height
    
    cmdDraw.Left = Me.ScaleWidth - 10 - cmdDraw.Width
    cmdDraw.Top = Me.ScaleHeight - 10 - StatusBar.Height - cmdDraw.Height
    
    Label1.Top = Me.ScaleHeight - 10 - StatusBar.Height - Label1.Height - 5
    
    Dim g As New Gradient
    g.Rectangle4Colors Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, vbWhite, vbWhite, RGB(224, 240, 255), RGB(245, 255, 224)
    'g.Rectangle Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, vbWhite, RGB(245, 255, 224), DiagonalFromBottom
    Me.ScaleMode = vbCentimeters
    picFunction.Cls
    
    DrawBorder picFunction
    DisplayAxisesAndGrid picFunction
    cmdDraw_Click
    
End Sub

Public Sub DisplayAxisesAndGrid(ByVal Object As Object)
    Dim Grid As LineStyle
    Dim Axises As LineStyle
    Dim NumsFont As New StdFont
    
    Grid.Visible = True
    Grid.DrawStyle = DrawStyleConstants.vbDot
    Grid.BorderWidth = 1
    Grid.BorderColor = RGB(128, 128, 128)
    
    Axises.Visible = True
    Axises.DrawStyle = DrawStyleConstants.vbSolid
    Axises.BorderColor = 0
    Axises.BorderWidth = 2
    
    NumsFont.Bold = True
    NumsFont.Name = "Tahoma"
    
    DrawAxisesAndGrid Object, MainAxisesStyle, MainGridStyle, MainNumbersFont, MainNumbersColor, MainNumbersVisible
         
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ms.Clear
    If IsSaved = False Or PromptSave = True Then
        Cancel = True
        Dim MsgBoxResult As VbMsgBoxResult
        MsgBoxResult = MsgBox("Save changes to '" & CurrentFileName & "' ?", vbYesNoCancel + vbQuestion)
        If MsgBoxResult = vbYes Then
            SaveAndUnload = False
            mnuFileSave_Click
            If SaveAndUnload Then
                ms.Clear
                UnloadForms
            End If
        ElseIf MsgBoxResult = vbNo Then
            Cancel = False
            ms.Clear
            UnloadForms
        ElseIf MsgBoxResult = vbCancel Then
        End If
    ElseIf IsSaved = True Then
        ms.Clear
        UnloadForms
    End If
End Sub
Private Sub UnloadForms()
    ms.Clear
    Unload frmAbout
    Unload frmAddNewConst
    Unload frmGridAndAxises
    Unload frmHelp
    Unload Me
    End
    'If App.LogMode <> 0 Then End
End Sub

Private Sub mnuFunctionSmoothDrawing_Click()
    mnuFunctionSmoothDrawing.Checked = Not (mnuFunctionSmoothDrawing.Checked)
    Form_Resize
End Sub

Private Sub mnuEditCopyFunctions_Click()
    Dim i As Integer
    Dim str As String
    For i = 0 To cmbFunction.ListCount - 1
        str = str & ProfessionalFormulaToLinear(cmbFunction.list(i)) & vbNewLine
    Next
    str = Left(str, Len(str) - 1)
    Clipboard.Clear
    Clipboard.SetText str
End Sub

Private Sub mnuEditCopyGraphImage_Click()
    Clipboard.Clear
    Clipboard.SetData picFunction.Image
End Sub

Private Sub mnuEditPasteFunctions_Click()

    If Clipboard.GetText = "" Then Exit Sub
    
    On Error Resume Next
    Dim i As Integer
    Dim StrArr() As String
    StrArr = Split(Clipboard.GetText, vbNewLine)
    For i = 0 To UBound(StrArr)
        cmbFunction.AddItem LinearFormulaToProfessional(StrArr(i))
    Next
    cmdDraw_Click

End Sub

Private Sub PasteFunctions()
    On Error Resume Next
    Dim i As Integer
    Dim StrArr() As String
    StrArr = Split(Clipboard.GetText, vbNewLine)
    For i = 0 To UBound(StrArr)
        cmbFunction.AddItem LinearFormulaToProfessional(StrArr(i))
    Next
    cmdDraw_Click
End Sub

Private Sub mnuFileClose_Click()
    Unload frmAbout
    Unload frmAddNewConst
    Unload frmGridAndAxises
    Unload frmHelp
    Unload frmMain
End Sub

Private Sub mnuFileNew_Click()
    Dim MsgBoxResult As VbMsgBoxResult
    If IsSaved = False Then
        MsgBoxResult = MsgBox("Save changes to '" & CurrentFileName & "' ?", vbYesNoCancel + vbQuestion)
        If MsgBoxResult = vbYes Then
            SaveMFDFileAs
            NewMFDFile
        ElseIf MsgBoxResult = vbNo Then
            NewMFDFile
        ElseIf MsgBoxResult = vbCancel Then
            Exit Sub
        End If
    ElseIf IsSaved = True Then
        NewMFDFile
    End If
    IsSaved = False
End Sub

Private Sub mnuFileOpen_Click()
    If IsSaved = False Or PromptSave = True Then
        Dim MsgBoxResult As VbMsgBoxResult
        MsgBoxResult = MsgBox("Save changes to '" & CurrentFileName & "' ?", vbYesNoCancel + vbQuestion)
        If MsgBoxResult = vbYes Then
            If IsSaved = False Then
                SaveMFDFileAs
            ElseIf PromptSave = True Then
                SaveMFDFile CurrentFileName
            Else
                SaveMFDFileAs
            End If
            OpenMFDFile
        ElseIf MsgBoxResult = vbNo Then
            OpenMFDFile
        ElseIf MsgBoxResult = vbCancel Then
        End If
    ElseIf IsSaved = True Then
        OpenMFDFile
    End If
End Sub

Private Sub mnuFilePrint_Click()
    pd.hwndOwner = Me.hwnd
    pd.MinAllowedPages = 0
    pd.MaxAllowedPages = 9999
    pd.ShowDialog
    
    'Printer.ScaleMode = vbPixels
    If pd.OK Then
        With Printer
            .ColorMode = pd.ColorMode
            .Copies = pd.Copies
            On Error Resume Next
            .Duplex = pd.DoubleSidedPrinting
            .Orientation = pd.Orientation
            .PaperBin = pd.PaperSource
            .PaperSize = pd.PaperSize
            .PrintQuality = pd.PrintQuality
            PrintPicture
        End With
    End If

End Sub

Private Sub mnuFileSave_Click()
    If IsSaved = False Then
        SaveMFDFileAs
    ElseIf IsSaved = True Then
        SaveMFDFile CurrentFilePath
    End If
End Sub

Public Sub mnuFileSaveAs_Click()
    SaveMFDFileAs
End Sub

Private Sub mnuFileSaveAsPicture_Click()
    Dim GdipSave As New GDIPlusPicture
    Dim SaveDlg As New SaveDialog
    
    With SaveDlg
        .AddExtension = True
        .DialogTitle = "Save Graph As Picture"
        .EnableSizing = True
        .Filter = "Bitmap Image(*.bmp)|*.bmp|JPEG Image(*.jpg)|*.jpg|GIF Graphics Interchange Format(*.gif)|*.gif|PNG Protable Network Graphics(*.png)|*.png|TIFF Tag Image File Format(*.tif)|*.tif|Widows MetaFile(*.wmf)|*.wmf|Enhanced Windows MetaFile(*.emf)|*.emf"
        .FilterIndex = 4
        .hwndOwner = Me.hwnd
        .InitialDirectory = GetPicturesFolder
        .InitialFileTitle = Left(CurrentFileName, Len(CurrentFileName) - 4)
        .PathNotExistWarning = True
        .ReplaceExistingFilePrompt = True
        .ShowDialog
        If .OK Then GdipSave.SavePictureFromHDC picFunction.hdc, picFunction.Picture, .FileName
    End With
End Sub

Private Sub mnuFunctionAddNewConstant_Click()
    frmAddNewConst.Form_Load
    frmAddNewConst.Show vbModal
End Sub

Private Sub mnuFunctionDrawTangent_Click()
    mnuOther.Caption = "Cancel"
    mnuOther.Visible = True
    DrawTangentMode = True
    InitMenu
End Sub

Private Sub mnuOther_Click()
    If mnuOther.Caption = "Cancel" Then
        DrawTangentMode = False
        mnuOther.Caption = ""
        mnuOther.Visible = False
        picFunction.Cls
        StatusBar.Panels(5).Text = ""
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Form_Load
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpContents_Click()
    frmHelp.Form_Load
    frmHelp.Show vbModal
End Sub

Private Sub mnuOptionsDrawStyles_Click(Index As Integer)
    PenHatchStyle = IndexToHatchStyle(Index)
    
    ReDim Preserve PenStyles(0 To cmbFunction.ListCount - 1)
    If cmbFunction.ListIndex = -1 Then cmbFunction.ListIndex = cmbFunction.ListCount - 1
    PenStyles(cmbFunction.ListIndex) = PenHatchStyle

    ClearMnuOptionsDrawStyle
    mnuOptionsDrawStyles(HatchStyleToIndex(PenStyles(cmbFunction.ListIndex))).Checked = True
    cmdDraw_Click
    PromptSave = True
End Sub

Private Sub mnuOptionsBorderWidths_Click(Index As Integer)
    On Error Resume Next
    ReDim Preserve BorderWidths(0 To cmbFunction.ListCount - 1)
    BorderWidths(UBound(BorderWidths)) = 1
    If cmbFunction.ListIndex = -1 Then cmbFunction.ListIndex = cmbFunction.ListCount - 1
    BorderWidths(cmbFunction.ListIndex) = IndexToBorderWidth(Index)
    
    ClearMnuOptionsBorderWidths
    mnuOptionsBorderWidths(BorderWidthToIndex(BorderWidths(cmbFunction.ListIndex))).Checked = True
    cmdDraw_Click
    PromptSave = True
End Sub

Private Sub mnuOptionsChangeColor_Click()

    ColorDlg.hwndOwner = Me.hwnd
    ColorDlg.InitializeRGB = True
    ColorDlg.CCDialogTitle = "Graph Line Color"
    ColorDlg.ShowDialog
    
    If ColorDlg.OK Then
        
        ReDim Preserve Colors(0 To cmbFunction.ListCount - 1)
        Colors(cmbFunction.ListIndex) = ColorDlg.Color
        
        ClearMnuOptionsDrawStyle
        mnuOptionsDrawStyles(HatchStyleToIndex(PenStyles(cmbFunction.ListIndex))).Checked = True
        cmdDraw_Click
    End If

End Sub

Private Function IndexToHatchStyle(ByVal Index As Integer) As Integer
    Select Case Index
    Case Is = 1
        IndexToHatchStyle = DottedSoft
    Case Is = 2
        IndexToHatchStyle = DottedMore
    Case Is = 3
        IndexToHatchStyle = Dashed
    Case Is = 4
        IndexToHatchStyle = NoHatch
    End Select
End Function

Private Function HatchStyleToIndex(ByVal HatchStyle As Integer) As Integer
    Select Case HatchStyle
    Case Is = DottedSoft
        HatchStyleToIndex = 1
    Case Is = DottedMore
        HatchStyleToIndex = 2
    Case Is = Dashed
        HatchStyleToIndex = 3
    Case Is = NoHatch
        HatchStyleToIndex = 4
    End Select
End Function

Private Function BorderWidthToIndex(ByVal BorderWidth As Single) As Integer
    Select Case BorderWidth
    Case Is = 2
        BorderWidthToIndex = 1
    Case Is = 3
        BorderWidthToIndex = 2
    Case Is = 4
        BorderWidthToIndex = 3
    Case Is = 5
        BorderWidthToIndex = 4
    Case Is = 6
        BorderWidthToIndex = 5
    Case Is = 7
        BorderWidthToIndex = 6
    Case Is = 10
        BorderWidthToIndex = 7
    Case Is = 15
        BorderWidthToIndex = 8
    End Select
End Function

Private Function IndexToBorderWidth(ByVal Index As Integer) As Single
    Select Case Index
    Case Is = 1
        IndexToBorderWidth = 2
    Case Is = 2
        IndexToBorderWidth = 3
    Case Is = 3
        IndexToBorderWidth = 4
    Case Is = 4
        IndexToBorderWidth = 5
    Case Is = 5
        IndexToBorderWidth = 6
    Case Is = 6
        IndexToBorderWidth = 7
    Case Is = 7
        IndexToBorderWidth = 10
    Case Is = 8
        IndexToBorderWidth = 15
    End Select
End Function

Private Sub mnuOptionsGrid_Click()
    With frmGridAndAxises
        Load frmGridAndAxises
        SaveAxisesStyle = MainAxisesStyle
        SaveGridStyle = MainGridStyle
        Set SaveNumbersFont = MainNumbersFont
        SaveNumbersColor = MainNumbersColor
        SaveNumbersVisible = MainNumbersVisible
        .Applied = True
        .Form_Load
        .Show vbModal
        
        .InitProperties
        '.chkAxisesVisible.Value = MainAxisesStyle.Visible
        '.cmbAxisesLineWidth.Text = MainAxisesStyle.BorderWidth
        '.picAxisesColor.BackColor = MainAxisesStyle.B
    End With
End Sub

Private Sub picFunction_DblClick()
    If DrawTangentMode = True Then
        cmbFunction.AddItem Trim(Replace(StatusBar.Panels(5).Text, "Tangent formula: ", ""))
        cmbFunction.ListIndex = cmbFunction.ListCount - 1
        DrawTangentMode = False
        cmdDraw_Click
        mnuOther.Caption = ""
        mnuOther.Visible = False
    End If
End Sub

Private Sub picFunction_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then Me.PopupMenu mnuOptions, , picFunction.Left, picFunction.Top
End Sub

Private Sub picFunction_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuOptions
    End If
End Sub

Private Sub picFunction_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Static chrToolTip As String
    If lastX <> X Or lastY <> Y Then
    
        Dim xCorrect As Single, yCorrect As Single
        Dim Fx As Single
        
        xCorrect = Round(X - picFunction.ScaleWidth / 2, 3)
        yCorrect = Round(picFunction.ScaleHeight / 2 - Y, 3)
        Fx = sc.Eval(Replace(LCase(ProfessionalFormulaToLinear(cmbFunction.Text)), "x", xCorrect))
        
        StatusBar.Panels(2).Text = "X = " & format(xCorrect, "00.000")
        StatusBar.Panels(3).Text = "Y = " & format(yCorrect, "00.000")
        If Err.Number = 0 Then StatusBar.Panels(4) = Chr$(131) & "(" & xCorrect & ")=" & format(Fx, "00.000")
        
        If Err.Number <> 0 Then
            Dim ErrorStr  As String
            If Err.Number = 11 Then ' Devision by zero
                ErrorStr = "number devided by zero."
            ElseIf Err.Number = 5 Then
                ErrorStr = "not a real number or number devided by zero."
            ElseIf Err.Number = 6 Then
                ErrorStr = "Very big number."
            End If
            StatusBar.Panels(4) = Chr$(131) & "(" & xCorrect & ")= " & ErrorStr
        End If
        Err.Clear
        
        If DrawTangentMode Then
            picFunction.Cls
            StatusBar.Panels(5).Text = "Tangent formula: " & CompleteStringWithChars(DrawTangent(picFunction, ConstsAdded, ProfessionalFormulaToLinear(cmbFunction.Text), X, 3, picFunction.ScaleWidth, picFunction.ScaleHeight, NewARGB(255, 0, 0, 0), 1, DrawStyleConstants.vbSolid, mnuFunctionSmoothDrawing.Checked), " ", 20)
            picFunction.Refresh
        End If
        
    End If
    lastX = X
    lastY = Y
End Sub
Private Function CompleteStringWithChars(ByVal str As String, ByVal char As String, ByVal NumChars As Integer) As String
    CompleteStringWithChars = str & String(NumChars - Len(str), char)
End Function
Private Sub cmbFunction_Change()
    TrimComboItems cmbFunction

    Dim SaveSelStart As Integer
    SaveSelStart = cmbFunction.SelStart
    
    cmbFunction.Text = LinearFormulaToProfessional(cmbFunction.Text)
    
    cmbFunction.SelStart = SaveSelStart
    
    If cmbFunction.ListCount = 0 Then
    
        mnuFileSave.Enabled = False
        mnuFileSaveAs.Enabled = False
        
        If Trim(cmbFunction.Text = "") Then
            cmdDraw.Enabled = False
            mnuOptionsChangeColor.Enabled = False
            mnuOptionsBorderWidth.Enabled = False
            mnuOptionsDrawStyle.Enabled = False
            mnuEditCopyFunctions.Enabled = False
            mnuFunctionDrawTangent.Enabled = False
        End If
    Else
    
        mnuFileSave.Enabled = True
        mnuFileSaveAs.Enabled = True
        
        cmdDraw.Enabled = True
        mnuOptionsChangeColor.Enabled = True
        mnuOptionsBorderWidth.Enabled = True
        mnuOptionsDrawStyle.Enabled = True
        mnuEditCopyFunctions.Enabled = True
        mnuFunctionDrawTangent.Enabled = True
    End If

    PromptSave = True
End Sub

Private Sub cmbFunction_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim SaveListIndex As Integer
    If KeyCode = vbKeyReturn Then
        cmdDraw_Click
    ElseIf KeyCode = vbKeyA Then
        If Shift = vbCtrlMask Then
            cmbFunction.SelStart = 0
            cmbFunction.SelLength = Len(cmbFunction.Text)
        End If
    ElseIf KeyCode = (vbKeyC Or KeyCode = vbKeyX) Then
        If Shift = vbCtrlMask Then
            Clipboard.Clear
            Clipboard.SetText ProfessionalFormulaToLinear(cmbFunction.Text)
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If cmbFunction.ListCount > 0 Then
            On Error Resume Next
            SaveListIndex = cmbFunction.ListIndex
            cmbFunction.RemoveItem cmbFunction.ListIndex
            If SaveListIndex > 0 Then cmbFunction.ListIndex = SaveListIndex - 1
            cmbFunction.Refresh
        End If
        
    End If
    
    If cmbFunction.ListCount = 0 And Trim(cmbFunction.Text = "") Then
        cmdDraw.Enabled = False
        mnuOptionsChangeColor.Enabled = False
        mnuOptionsBorderWidth.Enabled = False
        mnuOptionsDrawStyle.Enabled = False
        mnuEditCopyFunctions.Enabled = False
        mnuFunctionDrawTangent.Enabled = False
    Else
        cmdDraw.Enabled = True
        mnuOptionsChangeColor.Enabled = True
        mnuOptionsBorderWidth.Enabled = True
        mnuOptionsDrawStyle.Enabled = True
        mnuEditCopyFunctions.Enabled = True
        mnuFunctionDrawTangent.Enabled = True
    End If

End Sub

Private Sub ClearMnuOptionsDrawStyle()
    Dim i As Integer
    For i = 1 To mnuOptionsDrawStyles.UBound
        mnuOptionsDrawStyles(i).Checked = False
    Next
End Sub

Private Sub ClearMnuOptionsBorderWidths()
    Dim i As Integer
    For i = 1 To mnuOptionsBorderWidths.UBound
        mnuOptionsBorderWidths(i).Checked = False
    Next
End Sub

Private Sub tmrFunctionDetailsOnMenu_Timer()
    On Error Resume Next

    ClearMnuOptionsDrawStyle
    mnuOptionsDrawStyles(HatchStyleToIndex(PenStyles(cmbFunction.ListIndex))).Checked = True

    ClearMnuOptionsBorderWidths
    mnuOptionsBorderWidths(BorderWidthToIndex(BorderWidths(cmbFunction.ListIndex))).Checked = True

End Sub

Public Sub SaveMFDFileAs()
    Dim SaveDlg As New SaveDialog
    
    FileSaveCounter = FileSaveCounter + 1
    CurrentFileName = "Functions " & CStr(FileSaveCounter) & ".mfd"
    
    If cmbFunction.ListCount > 0 Then
            
        With SaveDlg
            .DialogTitle = "Save Functions File As"
            .EnableSizing = True
            .Filter = "Math Function Database File(*.mfd)|*.mfd"
            .hwndOwner = Me.hwnd
            .InitialDirectory = App.path & "\My Files\"
            .InitialFileTitle = CurrentFileName
            .PathNotExistWarning = True
            .AddExtension = True
            .ShowDialog
            If .OK Then
                SaveEncryptedFile WriteMFDFile, EncryptionCode2, .FileName
                CurrentFileName = .FileTitle
                CurrentFilePath = .PathName
                IsSaved = True
                SaveAndUnload = True
            ElseIf .Cancel Then
                FileSaveCounter = FileSaveCounter - 1
                SaveAndUnload = False
            End If
        End With
        
    End If
End Sub

Private Sub SaveMFDFile(ByVal StrFileName As String)
    
    If cmbFunction.ListCount > 0 Then
            
        SaveEncryptedFile WriteMFDFile, EncryptionCode2, StrFileName
        CurrentFileName = StrFileName
        IsSaved = True
        SaveAndUnload = True
    End If
End Sub

Private Sub OpenMFDFile()

    Dim OpenDlg As New OpenDialog
    
    With OpenDlg
        .EnableSizing = True
        .FileNotExistWarning = True
        .Filter = "Math Functions Database File(*.mfd)|*.mfd|All Files|*.*"
        .HideReadOnlyCheckBox = True
        .hwndOwner = Me.hwnd
        .InitialDirectory = App.path & "\My Files\"
        .PathNotExistWarning = True
        .PromptToCreateFile = False
        .ShowDialog
        
        If .OK Then
            ReadMFDFile .FileName
            IsSaved = True
            CurrentFileName = .FileTitle
        End If
    End With
DontOpen:
End Sub
Private Sub NewMFDFile()
    With MainAxisesStyle
        .BorderColor = 0
        .BorderWidth = 2
        .DrawStyle = vbSolid
        .Visible = True
    End With
    
    With MainGridStyle
        .BorderColor = RGB(128, 128, 128)
        .BorderWidth = 1
        .DrawStyle = vbDot
        .Visible = True
    End With
    
    With MainNumbersFont
        .Bold = True
        .Charset = 1
        .Italic = False
        .Name = "Tahoma"
        .Size = 8
        .Strikethrough = False
        .Underline = False
        .Weight = 700
    End With
    MainNumbersColor = vbBlue
    MainNumbersVisible = True
    
    cmbFunction.Clear
    cmbFunction.AddItem "X"
    cmbFunction.ListIndex = 0
    cmdDraw_Click
End Sub

Private Function WriteMFDFile() As String
    Dim i As Integer
    Dim StrToSave As String
    
    If cmbFunction.ListCount > 0 Then
    
        StrToSave = "Math Function Database File" & vbNewLine & _
        "Version 1.0" & vbNewLine & _
        "Axises:" & LineStyleToString(MainAxisesStyle) & vbNewLine & _
        "Grid:" & LineStyleToString(MainGridStyle) & vbNewLine & _
        "Numbers:" & FontOptionsToString(MainNumbersFont, MainNumbersColor, MainNumbersVisible) & vbNewLine & _
        ""
        For i = 0 To cmbFunction.ListCount - 1
            StrToSave = StrToSave & _
                        cmbFunction.list(i) & "," & _
                        str(Colors(i)) & "," & _
                        str(PenStyles(i)) & "," & _
                        str(BorderWidths(i)) & _
                        vbNewLine
        Next
        StrToSave = Left(StrToSave, Len(StrToSave) - 1)
    End If
    WriteMFDFile = StrToSave
End Function

Private Sub ReadMFDFile(ByVal StrFileName As String)
    Dim StrGetData As String
    Dim GraphOptions() As String
    Dim i As Integer
    
    On Error GoTo CantRead
    StrFileName = LoadEncryptedFile(EncryptionCode2, StrFileName)
    
    If GetLineFromString(StrFileName, 0) <> "Math Function Database File" Then
        MsgBox "This is not the matching file format.", vbCritical, "Error"
        Exit Sub
    End If
    
    If GetLineFromString(StrFileName, 1) <> "Version 1.0" Then
        MsgBox "File version mismatch", vbCritical, "Error"
        Exit Sub
    End If
    
    StrGetData = Replace(GetLineFromString(StrFileName, 2), "Axises:", "")
    MainAxisesStyle = StringToLineStyle(StrGetData)
    
    StrGetData = Replace(GetLineFromString(StrFileName, 3), "Grid:", "")
    MainGridStyle = StringToLineStyle(StrGetData)
    
    StrGetData = Replace(GetLineFromString(StrFileName, 4), "Numbers:", "")
    
    StringToStdFont StrGetData, MainNumbersFont
    MainNumbersColor = FontColorFromFontOptionsString(StrGetData)
    MainNumbersVisible = FontVisibleFromFontOptionsString(StrGetData)
    
    cmbFunction.Clear
    
    ReDim Colors(GetNumOfLines(StrFileName) - 5)
    ReDim PenStyles(GetNumOfLines(StrFileName) - 5)
    ReDim BorderWidths(GetNumOfLines(StrFileName) - 5)
    
    For i = 5 To GetNumOfLines(StrFileName) - 1
                        
        StrGetData = GetLineFromString(StrFileName, i)
        GraphOptions = Split(StrGetData, ",")
        
        cmbFunction.AddItem GraphOptions(0)
        Colors(i - 5) = CLng(GraphOptions(1))
        PenStyles(i - 5) = CInt(GraphOptions(2))
        BorderWidths(i - 5) = CSng(GraphOptions(3))
        
    Next
    cmdDraw_Click
    cmbFunction.ListIndex = 0
    Exit Sub
CantRead:
    MsgBox "An error occured while loading the file.", vbCritical
End Sub
