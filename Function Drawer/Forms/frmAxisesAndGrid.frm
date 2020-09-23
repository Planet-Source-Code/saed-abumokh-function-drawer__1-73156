VERSION 5.00
Begin VB.Form frmGridAndAxises 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Axises And Grid Options"
   ClientHeight    =   7440
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAxisesAndGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13.123
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   10.557
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPreview 
      Interval        =   1
      Left            =   2760
      Top             =   5160
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2130
      TabIndex        =   40
      Top             =   6990
      Width           =   1215
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   1770
      ScaleHeight     =   6.747
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   6.747
      TabIndex        =   36
      Top             =   2820
      Width           =   3855
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6885
      Left            =   60
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   2
      Top             =   60
      Width           =   5805
      Begin VB.CommandButton cmdNumbers 
         Caption         =   "Numbers"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrid 
         Caption         =   "Grid"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1215
      End
      Begin VB.CommandButton cmdAxises 
         Caption         =   "Axises"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox picNumbers 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   1530
         ScaleHeight     =   457
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   283
         TabIndex        =   12
         Top             =   0
         Width           =   4245
         Begin VB.CommandButton cmdNumberChoosesFont 
            Caption         =   "..."
            Height          =   360
            Left            =   1470
            TabIndex        =   35
            Top             =   1710
            Width           =   405
         End
         Begin VB.PictureBox picNumbersColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1470
            ScaleHeight     =   315
            ScaleWidth      =   705
            TabIndex        =   32
            Top             =   1170
            Width           =   705
         End
         Begin VB.CheckBox chkNumbersVisible 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show numbers on screen"
            Height          =   285
            Left            =   300
            TabIndex        =   31
            Top             =   750
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CommandButton cmdNumbersChooseColor 
            Caption         =   "..."
            Height          =   360
            Left            =   2250
            TabIndex        =   30
            Top             =   1155
            Width           =   405
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   37
            Top             =   2490
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numbers Font:"
            Height          =   195
            Left            =   300
            TabIndex        =   34
            Top             =   1830
            Width           =   1065
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numbers Color:"
            Height          =   195
            Left            =   300
            TabIndex        =   33
            Top             =   1260
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numbers"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1410
         End
      End
      Begin VB.PictureBox picGrid 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   1530
         ScaleHeight     =   457
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   283
         TabIndex        =   14
         Top             =   0
         Width           =   4245
         Begin VB.CommandButton cmdGridChooseColor 
            Caption         =   "..."
            Height          =   360
            Left            =   1980
            TabIndex        =   25
            Top             =   1605
            Width           =   405
         End
         Begin VB.ComboBox cmbGridLineStyle 
            Height          =   315
            ItemData        =   "frmAxisesAndGrid.frx":08CA
            Left            =   1200
            List            =   "frmAxisesAndGrid.frx":08DD
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1170
            Width           =   1755
         End
         Begin VB.CheckBox chkGridVisible 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show grid on screen"
            Height          =   285
            Left            =   300
            TabIndex        =   23
            Top             =   750
            Value           =   1  'Checked
            Width           =   2235
         End
         Begin VB.PictureBox picGridColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1200
            ScaleHeight     =   315
            ScaleWidth      =   705
            TabIndex        =   22
            Top             =   1620
            Width           =   705
         End
         Begin VB.ComboBox cmbGridLineWidth 
            Height          =   315
            ItemData        =   "frmAxisesAndGrid.frx":0912
            Left            =   1200
            List            =   "frmAxisesAndGrid.frx":0955
            TabIndex        =   21
            Text            =   "1"
            Top             =   2100
            Width           =   795
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   38
            Top             =   2490
            Width           =   675
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Line Syle:"
            Height          =   195
            Left            =   300
            TabIndex        =   29
            Top             =   1230
            Width           =   690
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Line Width:"
            Height          =   195
            Left            =   300
            TabIndex        =   28
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "points"
            Height          =   195
            Left            =   2100
            TabIndex        =   27
            Top             =   2160
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Line Color:"
            Height          =   195
            Left            =   300
            TabIndex        =   26
            Top             =   1710
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grid"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   660
         End
      End
      Begin VB.PictureBox picAxises 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   1530
         ScaleHeight     =   457
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   283
         TabIndex        =   6
         Top             =   0
         Width           =   4245
         Begin VB.ComboBox cmbAxisesLineWidth 
            Height          =   315
            ItemData        =   "frmAxisesAndGrid.frx":09AD
            Left            =   1200
            List            =   "frmAxisesAndGrid.frx":09F3
            TabIndex        =   18
            Text            =   "2"
            Top             =   2100
            Width           =   795
         End
         Begin VB.PictureBox picAxisesColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1200
            ScaleHeight     =   315
            ScaleWidth      =   705
            TabIndex        =   16
            Top             =   1620
            Width           =   705
         End
         Begin VB.CheckBox chkAxisesVisible 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show axises on screen"
            Height          =   285
            Left            =   300
            TabIndex        =   9
            Top             =   750
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.ComboBox cmbAxisesLineStyle 
            Height          =   315
            ItemData        =   "frmAxisesAndGrid.frx":0A4F
            Left            =   1200
            List            =   "frmAxisesAndGrid.frx":0A62
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1170
            Width           =   1755
         End
         Begin VB.CommandButton cmdAxisesChooseColor 
            Caption         =   "..."
            Height          =   360
            Left            =   1980
            TabIndex        =   7
            Top             =   1605
            Width           =   405
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   39
            Top             =   2490
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Line Color:"
            Height          =   195
            Left            =   300
            TabIndex        =   20
            Top             =   1710
            Width           =   765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "points"
            Height          =   195
            Left            =   2100
            TabIndex        =   19
            Top             =   2160
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Line Width:"
            Height          =   195
            Left            =   300
            TabIndex        =   17
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Axises"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Line Syle:"
            Height          =   195
            Left            =   300
            TabIndex        =   10
            Top             =   1230
            Width           =   690
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3390
      TabIndex        =   1
      Top             =   6990
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4650
      TabIndex        =   0
      Top             =   6990
      Width           =   1215
   End
End
Attribute VB_Name = "frmGridAndAxises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const BF_RIGHT = &H4

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Dim NumbersFont As New StdFont

Dim CustColors(1 To 16) As Long
Dim ColorDlg As New ColorDialog
Dim FontDlg As New FontDialog
Public Applied As Boolean

Private Sub chkAxisesVisible_Click()
    Select Case chkAxisesVisible.Value
    Case Is = vbChecked
        Label4.Enabled = True
        Label7.Enabled = True
        Label2.Enabled = True
        Label6.Enabled = True
        cmbAxisesLineStyle.Enabled = True
        cmbAxisesLineWidth.Enabled = True
        cmdAxisesChooseColor.Enabled = True
    Case Is = vbUnchecked
        Label4.Enabled = False
        Label7.Enabled = False
        Label2.Enabled = False
        Label6.Enabled = False
        cmbAxisesLineStyle.Enabled = False
        cmbAxisesLineWidth.Enabled = False
        cmdAxisesChooseColor.Enabled = False
    End Select
    Applied = False
End Sub

Private Sub chkGridVisible_Click()
    Select Case chkGridVisible.Value
    Case Is = vbChecked
        Label12.Enabled = True
        Label8.Enabled = True
        Label10.Enabled = True
        Label9.Enabled = True
        cmbGridLineStyle.Enabled = True
        cmbGridLineWidth.Enabled = True
        cmdGridChooseColor.Enabled = True
    Case Is = vbUnchecked
        Label12.Enabled = False
        Label8.Enabled = False
        Label10.Enabled = False
        Label9.Enabled = False
        cmbGridLineStyle.Enabled = False
        cmbGridLineWidth.Enabled = False
        cmdGridChooseColor.Enabled = False
    End Select
    Applied = False
End Sub

Private Sub chkNumbersVisible_Click()
    Select Case chkNumbersVisible.Value
    Case Is = vbChecked
        Label15.Enabled = True
        Label11.Enabled = True
        cmdNumbersChooseColor.Enabled = True
        cmdNumberChoosesFont.Enabled = True
    Case Is = vbUnchecked
        Label15.Enabled = False
        Label11.Enabled = False
        cmdNumbersChooseColor.Enabled = False
        cmdNumberChoosesFont.Enabled = False
    End Select
    Applied = False
End Sub

Private Sub cmbAxisesLineStyle_Change()
    Applied = False
End Sub

Private Sub cmbAxisesLineStyle_Click()
    Applied = False
End Sub

Private Sub cmbAxisesLineWidth_Change()
    Applied = False
End Sub

Private Sub cmbAxisesLineWidth_Click()
    Applied = False
End Sub

Private Sub cmbGridLineStyle_Change()
    Applied = False
End Sub

Private Sub cmbGridLineStyle_Click()
    Applied = False
End Sub

Private Sub cmbGridLineWidth_Change()
    Applied = False
End Sub

Private Sub cmbGridLineWidth_Click()
    Applied = False
End Sub

Private Sub cmdApply_Click()
    ChangeOptions
    frmMain.picFunction.Cls
    frmMain.picFunction.Picture = Nothing
    frmMain.DisplayAxisesAndGrid frmMain.picFunction
    frmMain.DrawFunctions
    Applied = True
End Sub

Private Sub cmdAxises_Click()
    picAxises.ZOrder
    UnboldBottons
    cmdAxises.Font.Bold = True
    SetAntiAliasedFontControls Me
End Sub

Private Sub cmdAxisesChooseColor_Click()
    ColorDlg.hwndOwner = Me.hWnd
    ColorDlg.InitializeRGB = True
    ColorDlg.CCDialogTitle = "Axises Color"
    ColorDlg.ShowDialog
    
    If ColorDlg.OK Then
        picAxisesColor.BackColor = ColorDlg.Color
        Applied = False
    End If

End Sub

Private Sub cmdCancel_Click()
    ResetsOptions
    frmMain.picFunction.Cls
    frmMain.DisplayAxisesAndGrid frmMain.picFunction
    frmMain.DrawFunctions
    Me.Hide
End Sub

Private Sub cmdGrid_Click()
    picGrid.ZOrder
    UnboldBottons
    cmdGrid.Font.Bold = True
    SetAntiAliasedFontControls Me
End Sub

Private Sub cmdGridChooseColor_Click()
    ColorDlg.hwndOwner = Me.hWnd
    ColorDlg.InitializeRGB = True
    ColorDlg.CCDialogTitle = "Grid Color"
    ColorDlg.ShowDialog
    
    If ColorDlg.OK Then
        picGridColor.BackColor = ColorDlg.Color
        Applied = False
    End If
End Sub

Private Sub cmdNumberChoosesFont_Click()

    FontDlg.hwndOwner = Me.hWnd
    FontDlg.DialogTitle = "Numbers Font"
    FontDlg.ShowScreenFonts = True
    FontDlg.ShowPrinterFonts = True
    FontDlg.InitializeFontProperties = True
    FontDlg.FontNotExistPrompt = True
    FontDlg.SpecifyCharsets = True
    FontDlg.FontCharSet = CharSets.DefaultCharSet
    
    FontDlg.FontName = NumbersFont.Name
    FontDlg.FontItalic = NumbersFont.Italic
    FontDlg.FontBold = NumbersFont.Bold
    FontDlg.FontSize = NumbersFont.Size
    
    FontDlg.ShowDialog
    
    If FontDlg.OK Then
        With NumbersFont
            .Bold = FontDlg.FontBold
            .Charset = FontDlg.FontCharSet
            .Italic = FontDlg.FontItalic
            .Name = FontDlg.FontName
            .Size = FontDlg.FontSize
        End With
        Applied = False
    End If
End Sub

Private Sub cmdNumbers_Click()
    picNumbers.ZOrder
    picNumbers.ZOrder
    UnboldBottons
    cmdNumbers.Font.Bold = True
    SetAntiAliasedFontControls Me
End Sub

Private Sub cmdNumbersChooseColor_Click()
    ColorDlg.hwndOwner = Me.hWnd
    ColorDlg.InitializeRGB = True
    ColorDlg.CCDialogTitle = "Numbers Color"
    ColorDlg.ShowDialog
    
    If ColorDlg.OK Then
        picNumbersColor.BackColor = ColorDlg.Color
        Applied = False
    End If

End Sub

Private Sub cmdOK_Click()
    If Applied = False Then cmdApply_Click
    frmMain.IsSaved = False
    Me.Hide
End Sub

Public Sub Form_Load()
    Dim r As RECT
    r.Left = cmdAxises.Left + cmdAxises.Width + 10
    r.Right = cmdAxises.Left + cmdAxises.Width + 10
    r.Top = 0
    r.Bottom = picMain.ScaleHeight
    
    DrawEdge picMain.hdc, r, EDGE_ETCHED, BF_RIGHT
    picMain.Refresh
    
    InitProperties
    
    cmdAxises_Click
    
    
    Dim i As Integer
    For i = 1 To 16
        CustColors(i) = QBColor(i - 1)
    Next
    ColorDlg.SetCustomColors CustColors
    ColorDlg.Color = RGB(128, 128, 128)

    Dim g As New Gradient
    g.Rectangle picMain.hdc, 0, 0, cmdAxises.Left + cmdAxises.Width + 10, picMain.ScaleHeight, RGB(192, 224, 255), RGB(255, 255, 255), False
    g.Rectangle4Colors picAxises.hdc, picAxises.ScaleWidth * 0.25, 0, picAxises.ScaleWidth, picAxises.ScaleHeight, vbWhite, vbWhite, vbWhite, RGB(255, 255, 224)
    g.Rectangle4Colors picGrid.hdc, picGrid.ScaleWidth * 0.25, 0, picGrid.ScaleWidth, picGrid.ScaleHeight, vbWhite, vbWhite, vbWhite, RGB(255, 255, 224)
    g.Rectangle4Colors picNumbers.hdc, picNumbers.ScaleWidth * 0.25, 0, picNumbers.ScaleWidth, picNumbers.ScaleHeight, vbWhite, vbWhite, vbWhite, RGB(255, 255, 224)
    'SetAntiAliasedLabels Me, picAxises
    'SetAntiAliasedLabels Me, picGrid
    'SetAntiAliasedLabels Me, picNumbers
End Sub

Private Sub UnboldBottons()
    cmdAxises.Font.Bold = False
    cmdGrid.Font.Bold = False
    cmdNumbers.Font.Bold = False
End Sub

Private Sub tmrPreview_Timer()

    Dim AxisesStyle As LineStyle
    Dim GridStyle As LineStyle

    AxisesStyle.BorderColor = picAxisesColor.BackColor
    AxisesStyle.BorderWidth = cmbAxisesLineWidth.Text
    AxisesStyle.DrawStyle = cmbAxisesLineStyle.ListIndex
    AxisesStyle.Visible = chkAxisesVisible.Value

    GridStyle.BorderColor = picGridColor.BackColor
    GridStyle.BorderWidth = Val(cmbGridLineWidth.Text)
    GridStyle.DrawStyle = cmbGridLineStyle.ListIndex
    GridStyle.Visible = chkGridVisible.Value


    picPreview.Cls
    DrawAxisesAndGrid picPreview, AxisesStyle, GridStyle, NumbersFont, picNumbersColor.BackColor, chkNumbersVisible.Value
    picPreview.Refresh
    If Applied = True Then
        cmdApply.Enabled = False
    Else
        cmdApply.Enabled = True
    End If
End Sub

Private Sub ChangeOptions()
    
    With MainAxisesStyle
        .BorderColor = picAxisesColor.BackColor
        .BorderWidth = Val(cmbAxisesLineWidth.Text)
        .DrawStyle = cmbAxisesLineStyle.ListIndex
        .Visible = chkAxisesVisible.Value
    End With
    
    With MainGridStyle
        .BorderColor = picGridColor.BackColor
        .BorderWidth = Val(cmbGridLineWidth.Text)
        .DrawStyle = cmbGridLineStyle.ListIndex
        .Visible = chkGridVisible.Value
    End With
    
    Set MainNumbersFont = NumbersFont
    
    MainNumbersColor = picNumbersColor.BackColor
    MainNumbersVisible = chkNumbersVisible.Value

End Sub

Private Sub ResetsOptions()
    MainAxisesStyle = SaveAxisesStyle
    MainGridStyle = SaveGridStyle
    Set MainNumbersFont = SaveNumbersFont
    MainNumbersColor = SaveNumbersColor
    MainNumbersVisible = SaveNumbersVisible
End Sub

Public Sub InitProperties()
    With MainAxisesStyle
        chkAxisesVisible.Value = Abs(.Visible)
        cmbAxisesLineWidth.Text = .BorderWidth
        picAxisesColor.BackColor = .BorderColor
        cmbAxisesLineStyle.ListIndex = .DrawStyle
    End With

    With MainGridStyle
        chkGridVisible.Value = Abs(.Visible)
        cmbGridLineWidth.Text = .BorderWidth
        picGridColor.BackColor = .BorderColor
        cmbGridLineStyle.ListIndex = .DrawStyle
    End With
    
    Set NumbersFont = MainNumbersFont
    chkNumbersVisible.Value = Abs(MainNumbersVisible)
    picNumbersColor.BackColor = MainNumbersColor
End Sub
