VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help Contents"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   375
      Left            =   6780
      TabIndex        =   2
      Top             =   5820
      Width           =   1245
   End
   Begin VB.PictureBox picIndex 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   90
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   90
      Width           =   2565
      Begin VB.VScrollBar vsIndex 
         Height          =   5655
         LargeChange     =   10
         Left            =   2250
         TabIndex        =   3
         Top             =   0
         Width           =   285
      End
      Begin VB.PictureBox picIndexScroll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8775
         Left            =   -30
         ScaleHeight     =   585
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   169
         TabIndex        =   4
         Top             =   -3120
         Width           =   2535
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Draw Tangent"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   17
            Left            =   600
            TabIndex        =   30
            Top             =   5040
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Function Tools"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   4740
            Width           =   1785
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Logical operators"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   11
            Left            =   300
            TabIndex        =   28
            Top             =   840
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Function Graphics"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   13
            Left            =   600
            TabIndex        =   26
            Top             =   6930
            Width           =   1485
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Draw More Than One Function"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Index           =   12
            Left            =   600
            TabIndex        =   25
            Top             =   6420
            Width           =   1755
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiple Function Drawing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   120
            TabIndex        =   24
            Top             =   5940
            Width           =   1455
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Other Functions and Operators"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Index           =   9
            Left            =   570
            TabIndex        =   23
            Top             =   4260
            Width           =   1890
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Inverse Hyperbolic Functions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Index           =   8
            Left            =   600
            TabIndex        =   22
            Top             =   3780
            Width           =   1395
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Hyperbolic Functions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   7
            Left            =   600
            TabIndex        =   21
            Top             =   3510
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Inverse Trigonometric Functions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Index           =   6
            Left            =   600
            TabIndex        =   20
            Top             =   3030
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "How to Write Formulas?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   90
            Width           =   2010
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Basic Operators"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   18
            Top             =   360
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Exponentiation and Roots"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   17
            Top             =   600
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Functions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   2010
            Width           =   810
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Logarithms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   600
            TabIndex        =   15
            Top             =   2250
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Trigonometric Functions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Index           =   5
            Left            =   600
            TabIndex        =   14
            Top             =   2550
            Width           =   1020
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Constants"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   1170
            Width           =   855
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Write a Constant"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   12
            Top             =   1410
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Add New Constant"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   11
            Top             =   1680
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Function Information"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   5370
            Width           =   1785
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Status Bar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   10
            Left            =   600
            TabIndex        =   9
            Top             =   5640
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Load and Save to File"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   14
            Left            =   600
            TabIndex        =   8
            Top             =   7350
            Width           =   1545
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   15
            Left            =   600
            TabIndex        =   7
            Top             =   7680
            Width           =   405
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Grid Options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   6
            Top             =   8160
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Grid And Axises Styles"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   16
            Left            =   600
            TabIndex        =   5
            Top             =   8430
            Width           =   1635
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.PictureBox picContent 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   2730
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   1
      Top             =   90
      Width           =   5265
      Begin RichTextLib.RichTextBox rtfHelp 
         Height          =   5235
         Left            =   210
         TabIndex        =   27
         Top             =   210
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   9234
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmHelp.frx":08CA
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lblIndex As Integer
Option Explicit
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Public FileLoaded As String
Private Sub cmdOK_Click()
    Me.Hide
End Sub

Public Sub Form_Load()
    picIndexScroll.Top = 0
    vsIndex.Max = picIndexScroll.Height - picIndex.Height
    vsIndex.LargeChange = picIndexScroll.Height
    Dim g As New Gradient
    g.Rectangle picIndexScroll.hdc, 0, 0, 10, picIndexScroll.ScaleHeight, vbWhite, RGB(192, 224, 255), False
    g.Rectangle picIndexScroll.hdc, 10, 0, picIndexScroll.ScaleWidth / 2, picIndexScroll.ScaleHeight, RGB(192, 224, 255), RGB(224, 240, 255), False
    g.Rectangle picIndexScroll.hdc, picIndexScroll.ScaleWidth / 2, 0, picIndexScroll.ScaleWidth, picIndexScroll.ScaleHeight, RGB(224, 240, 255), RGB(255, 255, 255), False
    g.Triangle picContent.hdc, 0, 0, vbWhite, picContent.ScaleWidth * 1.5, 0, RGB(200, 224, 255), 0, picContent.ScaleHeight * 1.5, RGB(200, 224, 255)
    g.Triangle picContent.hdc, picContent.ScaleWidth * 1.5, 0, RGB(200, 224, 255), 0, picContent.ScaleHeight * 1.5, RGB(200, 224, 255), picContent.ScaleWidth, picContent.ScaleHeight, vbWhite
    
    MouseWheel.Hook frmHelp.picIndex.hwnd
    
    FileLoaded = Replace(App.path & "\", "\\", "\") & "Data\Help\overview.rtf"
    rtfHelp.LoadFile FileLoaded
End Sub

Public Sub Label_Click(Index As Integer)
    lblIndex = Index
    FileLoaded = Replace(App.path & "\", "\\", "\") & "Data\Help\index " & Index & ".rtf"
    rtfHelp.LoadFile FileLoaded
End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To Label.UBound
        Label(i).ForeColor = vbBlue
    Next
    Label(Index).ForeColor = vbRed
End Sub

Private Sub picIndexScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To Label.UBound
        Label(i).ForeColor = vbBlue
    Next
End Sub

Private Sub rtfHelp_Change()
    Dim SaveSelStart As Long
    SaveSelStart = rtfHelp.SelStart
    rtfHelp.LoadFile FileLoaded
    rtfHelp.SelStart = SaveSelStart
End Sub

Private Sub rtfHelp_Click()
    HideCaret rtfHelp.hwnd
End Sub

Private Sub rtfHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideCaret rtfHelp.hwnd
End Sub

Private Sub rtfHelp_SelChange()
    HideCaret rtfHelp.hwnd
End Sub

Private Sub vsIndex_Change()
    picIndexScroll.Top = -vsIndex.Value
End Sub

Private Sub vsIndex_GotFocus()
    picIndex.SetFocus
End Sub

Private Sub vsIndex_Scroll()
    vsIndex_Change
End Sub

Private Function IsInRect(ByVal X As Long, ByVal Y As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As Boolean
    If (X >= Left And X <= (Left + Width)) And (Y >= Top And Y <= (Top + Height)) Then IsInRect = True
End Function

Public Function Mouse_Wheel(ByVal WheelDelta As Integer, ByVal Button As Integer, ByVal X As Integer, ByVal Y As Integer)
    
    On Error Resume Next
    
    Dim Delta As Long
    Dim i As Integer
    
    Delta = -WheelDelta / 120 * Label(0).Height * 3
    
    Do Until i = Delta
        DoEvents
        vsIndex.Value = vsIndex.Value + Sgn(Delta)
        i = i + Sgn(Delta)
    Loop

End Function
