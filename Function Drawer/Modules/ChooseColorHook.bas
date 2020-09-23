Attribute VB_Name = "CommonDialogsHooks"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or diStribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const WM_INITDIALOG As Long = &H110

'ChooseColor Dialog component control ID's
'dialog buttons
Private Const CTLID_BTN_OK  As Long = 1
Private Const CTLID_BTN_CANCEL  As Long = 2

'for color dialog
Private Const CTLID_BTN_ADDTOCUSTOMCOLORS As Long = &H2C8
Private Const CTLID_BTN_DEFINECUSTOMCOLORS As Long = &H2CF

'for font dialog
Private Const CTLID_BTN_APPLY As Long = &H402
Private Const CTLID_COMBO_COLOR As Long = &H473

'labels
Private Const CTLID_LABEL_HUE As Long = &H2D3
Private Const CTLID_LABEL_SAT As Long = &H2D4
Private Const CTLID_LABEL_LUM As Long = &H2D5
Private Const CTLID_LABEL_RED As Long = &H2D6
Private Const CTLID_LABEL_BLUE As Long = &H2D7
Private Const CTLID_LABEL_GREEN As Long = &H2D8

'text boxes
Private Const CTLID_VALUE_HUE As Long = &H2BF
Private Const CTLID_VALUE_SAT As Long = &H2C0
Private Const CTLID_VALUE_LUM As Long = &H2C1
Private Const CTLID_VALUE_RED As Long = &H2C2
Private Const CTLID_VALUE_BLUE As Long = &H2C3
Private Const CTLID_VALUE_GREENE As Long = &H2C4

'palettes / selectors
Private Const CTLID_PALETTE_BASICCOLORS As Long = &H2D0
Private Const CTLID_PALETTE_CUSTOM_COLORS As Long = &H2D1
Private Const CTLID_PALETTE_CUSTOMRAINBOW As Long = &H2C6
Private Const CTLID_PALETTE_CUSTOMDENSITY As Long = &H2BE

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Public CCDialogTitle As String
Public CCOKButtonCaption As String
Public CCCancelButtonCaption As String
Public CCDefineCustomColorsButtonCaption As String
Public CCAddToCustomColorsButtonCaption As String

Public CFDialogTitle As String
Public CFOKButtonCaption As String
Public CFCancelButtonCaption As String
Public CFEnableColorComboBox As Boolean

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Function ChooseColorProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   Dim rc          As RECT
   Dim hWndControl As Long
   Dim scrWidth    As Long
   Dim scrHeight   As Long
   Dim dlgWidth    As Long
   Dim dlgHeight   As Long

   Select Case uMsg
      Case WM_INITDIALOG
      
        Call SetWindowText(hWnd, CCDialogTitle)
        
        hWndControl = GetDlgItem(hWnd, CTLID_BTN_OK)
        SetWindowText hWndControl, CCOKButtonCaption
        
        hWndControl = GetDlgItem(hWnd, CTLID_BTN_CANCEL)
        SetWindowText hWndControl, CCCancelButtonCaption
        
        hWndControl = GetDlgItem(hWnd, CTLID_BTN_ADDTOCUSTOMCOLORS)
        SetWindowText hWndControl, CCAddToCustomColorsButtonCaption
        
        hWndControl = GetDlgItem(hWnd, CTLID_BTN_DEFINECUSTOMCOLORS)
        SetWindowText hWndControl, CCDefineCustomColorsButtonCaption
        
      Case Else
   End Select
End Function

Public Function ChooseFontProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim rc          As RECT
   Dim hWndControl As Long
   Dim scrWidth    As Long
   Dim scrHeight   As Long
   Dim dlgWidth    As Long
   Dim dlgHeight   As Long

   Select Case uMsg
      Case WM_INITDIALOG
      
         Call SetWindowText(hWnd, CFDialogTitle)
         
        hWndControl = GetDlgItem(hWnd, CTLID_BTN_OK)
        SetWindowText hWndControl, CFOKButtonCaption
        
        hWndControl = GetDlgItem(hWnd, CTLID_BTN_CANCEL)
        SetWindowText hWndControl, CFCancelButtonCaption
        
        hWndControl = GetDlgItem(hWnd, CTLID_COMBO_COLOR)
        EnableWindow hWndControl, CFEnableColorComboBox
        
      Case Else
   End Select
End Function
Public Function GetProc(ByVal dwProc As Long) As Long
   GetProc = dwProc
End Function


