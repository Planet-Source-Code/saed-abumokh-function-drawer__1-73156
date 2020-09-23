VERSION 5.00
Begin VB.Form frmAddNewConst 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Constant"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   Icon            =   "frmAddNewConst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtConstValue 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   990
      Width           =   3975
   End
   Begin VB.TextBox txtConstName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2790
      TabIndex        =   1
      Top             =   1380
      Width           =   1245
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   1380
      Width           =   1245
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Constant Value:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   750
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Constant Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   1350
   End
End
Attribute VB_Name = "frmAddNewConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim sc As New ScriptControl
    Set sc = frmMain.sc
    Dim Expression As String
    Dim errConst As String
    Dim i As Integer
    Dim char As String
    
    On Error Resume Next
    
    Expression = "const " & txtConstName & " = " & sc.Eval(txtConstValue.Text) & vbNewLine
    
    
    For i = 1 To Len(txtConstValue.Text)
        char = Strings.Mid$(txtConstValue.Text, i, 1)
        If char = "=" Or char = ">" Or char = "<" Then
            MsgBox "Constant value must not include comparison signs."
            Exit Sub
        End If
    Next
        
    sc.AddCode Expression
    errConst = Err.description
    
    If Err = 0 Then
        Me.Hide
        With frmMain
            .ConstsAdded = .ConstsAdded & Expression & vbNewLine
        End With
    Else
        If Err.Number = 1010 Then 'expected identifier
            errConst = "Write a constant name please."
        ElseIf Err.Number = 1011 Then ' expected '='
            errConst = "Write a valid constant name please."
        ElseIf Err.Number = 1041 Then ' expected '='
            errConst = "This constant is already defined."
        ElseIf Err.Number = 1002 Then ' syntax error
            If txtConstValue.Text = "" Then
                errConst = "Write the constant value please."
            Else
                errConst = "Constant value syntax error."
            End If
        ElseIf Err.Number = 1032 Then ' syntax error
            errConst = "Please write valid characters."
        ElseIf Err.Number = 1045 Then ' expected literal constant
            errConst = "Constant name must not include comparison signs."
        ElseIf Err.Number = 6 Then ' over flow
            errConst = "The constant value is too big."
        Else
            errConst = "Unexpected error, Cannot add the contant."
        End If
        MsgBox errConst & " " & Err.Number
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdAdd_Click
End Sub

Public Sub Form_Load()
    Dim g As New Gradient
    g.Rectangle Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdAdd.Height - 20, RGB(224, 240, 255), RGB(255, 255, 255), True
    g.Rectangle Me.hdc, 0, Me.ScaleHeight - cmdAdd.Height - 20, Me.ScaleWidth, Me.ScaleHeight, RGB(255, 255, 255), RGB(224, 240, 255), True
    SetAntiAliasedLabels Me, Me
End Sub
