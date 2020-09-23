VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Friend Property Get ImageFiles() As String()
    ImageFiles = msImageFiles
End Property

Friend Property Let ImageFiles(ByRef vNewValue() As String)
   msImageFiles = vNewValue
End Property

Friend Property Get ImageWidth() As Long
    ImageWidth = msImageWidth
End Property

Friend Property Let ImageWidth(ByVal vNewValue As Long)
   msImageWidth = vNewValue
End Property

Friend Property Get ImageHeight() As Long
    ImageHeight = msImageHeight
End Property

Friend Property Let ImageHeight(ByVal vNewValue As Long)
   msImageHeight = vNewValue
End Property

Friend Property Get MenuFont() As StdFont
    MenuFont = msMenuFont
End Property

Friend Property Let MenuFont(ByVal vNewValue As StdFont)
   msMenuFont = vNewValue
End Property

Friend Property Get hWnd() As Long
    hWnd = mshWnd
End Property

Friend Property Let hWnd(ByVal vNewValue As Long)
   mshWnd = vNewValue
End Property

Public Sub SetStyle()
    HookMenu mshWnd
End Sub

Public Property Get MenuCount() As Long
    MenuCount = GetMenuCount
End Property

Public Sub Clear()
    ClearAll mshWnd
End Sub
Private Sub Class_Initialize()
    msImageWidth = 16
    msImageHeight = 16
End Sub
