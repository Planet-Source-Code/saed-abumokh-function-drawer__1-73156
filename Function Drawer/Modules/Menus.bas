Attribute VB_Name = "Menus"
Option Explicit

Private Const MF_BYPOSITION = &H400&
Private Const MIM_BACKGROUND = &H2

Private Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack  As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hwnd As Long, mInfo As MENUINFO) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Sub SetSubMenuItemBitmap(ByVal hwnd As Long, ByVal MenuPosition As Long, ByVal SubMenuPosition As Long, ByVal Bitmap As Long)
    Dim hMenu As Long, hSubMenu As Long
    hMenu = GetMenu(hwnd)
    hSubMenu = GetSubMenu(hMenu, MenuPosition)
    SetMenuItemBitmaps hSubMenu, SubMenuPosition, MF_BYPOSITION, Bitmap, Bitmap
End Sub

Public Sub SetSubSubMenuItemBitmap(ByVal hwnd As Long, ByVal MenuPosition As Long, ByVal SubMenuPosition As Long, ByVal SubSubMenuPosition As Long, ByVal Bitmap As Long)
    Dim hMenu As Long, hSubMenu As Long, hSubSubMenu As Long
    hMenu = GetMenu(hwnd)
    hSubMenu = GetSubMenu(hMenu, MenuPosition)
    hSubSubMenu = GetSubMenu(hSubMenu, SubMenuPosition)
    SetMenuItemBitmaps hSubSubMenu, SubSubMenuPosition, MF_BYPOSITION, Bitmap, Bitmap
End Sub

Public Sub SetMenuBarBackground(ByVal hwnd As Long, ByVal hBitmap As Long)
    Dim hMenuInfo As MENUINFO
    hMenuInfo.cbSize = Len(hMenuInfo)
    hMenuInfo.fMask = MIM_BACKGROUND
    hMenuInfo.hbrBack = CreatePatternBrush(hBitmap)
    SetMenuInfo GetMenu(hwnd), hMenuInfo
    DrawMenuBar hwnd
End Sub

Public Sub SetMenuBackground(ByVal hwnd As Long, ByVal MenuPosition As Long, ByVal hBitmap As Long)
    Dim hMenuInfo As MENUINFO
    hMenuInfo.cbSize = Len(hMenuInfo)
    hMenuInfo.fMask = MIM_BACKGROUND
    hMenuInfo.hbrBack = CreatePatternBrush(hBitmap)
    SetMenuInfo GetSubMenu(GetMenu(hwnd), MenuPosition), hMenuInfo
End Sub

