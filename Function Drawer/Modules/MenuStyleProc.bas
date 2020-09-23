Attribute VB_Name = "MenuStyleProc"
Option Explicit
'GetWindowLong and proc
Private Const GWL_WNDPROC = (-4)
Dim PrevWndProc As Long

'Window Messages: wm_
Private Const WM_MENUSELECT = &H11F
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const WM_CLOSE = &H10

'menu info masks: mim_
Private Const MIM_BACKGROUND = &H2

'menu item info masks: miim_
Private Const MIIM_STRING = &H40
Private Const MIIM_FTYPE = &H100
Private Const MIIM_TYPE = &H10

'menu fTypes: mft_
Private Const MFT_STRING = &H0

'owner draw types: odt_
Private Const ODT_MENU = 1

'owner draw states:
Private Const ODS_SELECTED = &H1
Private Const ODS_GRAYED = &H2
Private Const ODS_DISABLED = &H4
Private Const ODS_CHECKED = &H8
Private Const ODS_FOCUS = &H10
Private Const ODS_HOTLIGHT = &H40
Private Const ODS_INACTIVE = &H80
Private Const ODS_NOACCEL = &H100
Private Const ODS_NOFOCUSRECT = &H200

'Menu Flags: mf_
Private Const MF_INSERT = &H0&
Private Const MF_CHANGE = &H80&
Private Const MF_APPEND = &H100&
Private Const MF_DELETE = &H200&
Private Const MF_REMOVE = &H1000&

Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&

Private Const MF_SEPARATOR = &H800&

Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_DISABLED = &H2&

Private Const MF_UNCHECKED = &H0&
Private Const MF_CHECKED = &H8&
Private Const MF_USECHECKBITMAPS = &H200&

Private Const MF_STRING = &H0&
Private Const MF_BITMAP = &H4&
Private Const MF_OWNERDRAW = &H100&

Private Const MF_POPUP = &H10&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&

Private Const MF_UNHILITE = &H0&
Private Const MF_HILITE = &H80&

Private Const MF_SYSMENU = &H2000&
Private Const MF_HELP = &H4000&
Private Const MF_MOUSESELECT = &H8000&

'DrawEdge borders: bdr_ (not for use)
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

'DrawEdge edges (for use)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

'DrawEdge border flags (rectangle sides): bf_
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

'DrawEdge border flags (rectangle type): bf_
Private Const BF_MIDDLE = &H800     ' Fill in the middle.
Private Const BF_SOFT = &H1000     ' Use for softer buttons.
Private Const BF_ADJUST = &H2000   ' Calculate the space left over.
Private Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Private Const BF_MONO = &H8000     ' For monochrome borders.

Private Const LOGPIXELSY = 90         '  Logical pixels/inch in Y for CreateFont
Private Const PS_NULL = 5  'to remove shape border

'to get text width and height using GetTextExtentPoint32
Private Type Size
        cx As Long
        cy As Long
End Type

'to draw the sub menu pointer triangle
Private Type POINTAPI
        X As Long
        Y As Long
End Type

'for getting bitmap size
Private Type Bitmap '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

'to get menu item area when drawind on it
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'to draw custom menus
Private Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        itemData As Long
End Type

'to change menu item size
Private Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        itemHeight As Long
        itemData As Long
End Type

' to change menu item properties such as owner draw property
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

'to set menu and menu bar background
Private Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack  As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

'menu APIs
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hwnd As Long, mInfo As MENUINFO) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long ' refresh menu bar
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long

'Events proc APIs
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'to fill the structures from lParam in CallWindowProc
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Graphics APIs
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

'APIs  font width and height
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal Height As Long, ByVal Width As Long, ByVal Escapement As Long, ByVal Orientation As Long, ByVal Weight As Long, ByVal Italic As Long, ByVal Underline As Long, ByVal StrikeOut As Long, ByVal Charset As Long, ByVal OutPrecision As Long, ByVal ClipPrecision As Long, ByVal Quality As Long, ByVal PitchAndFamily As Long, ByVal FaceName As String) As Long

'GDI objects font APIs
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


'owner drawn menu properties:
'Public Captions() As String ' array of captions of menus (cannot get caption by menu item using API)
'Public HotKeys() As String 'array of hot keys (key accelators) (cannot get them by menu item using API)
Dim ImageFiles() As String ' array of image files of menus ,not handles, that we can
                              ' load cool images with alpha pixels using GDI+
Private Type MenuType
    MenuIndex As Long
    MenuHandle As Long
    MenuText As String
    HotKey As String
    Status As MenuStatus
End Type

Private Enum MenuStatus
    NormalItem = 0
    Separator = 1
    HasSubs = 2
End Enum

Dim MyMenus() As MenuType

Public msImageFiles() As String
Public msImageWidth As Long
Public msImageHeight As Long
Public msMenuFont As New StdFont
Public mshWnd As Long

Dim gp As New GDIPlusPicture
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Public Sub HookMenu(ByVal hwnd As Long)
    GetMenuStrings GetMenu(hwnd), MyMenus
    ReDim Preserve msImageFiles(UBound(MyMenus)) As String
    SetAllMenusOwnerDraw GetMenu(hwnd)
    PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MenuProc)
    
End Sub

Public Sub ClearAll(ByVal hwnd As Long)
    SetWindowLong hwnd, GWL_WNDPROC, PrevWndProc
    RemoveAllMenusOwnerDraw GetMenu(hwnd)
    DrawMenuBar hwnd
End Sub

Private Function GetTextSize(ByVal str As String, ByVal Font As StdFont) As Size

    Dim dc As New DeviceContext
    Dim s As Size
    Dim hFont As Long, OldFont As Long
    
    dc.Create 24, 0, 0
    hFont = CreateFont(-MulDiv(Font.Size, GetDeviceCaps(dc.Handle, LOGPIXELSY), 72), 0, 0, 0, Font.Weight, Font.Italic, Font.Underline, Font.Strikethrough, Font.Charset, 0, 0, 0, 0, Font.Name)
    OldFont = SelectObject(dc.Handle, hFont)
    GetTextExtentPoint32 dc.Handle, str, Len(str), s
    SelectObject dc.Handle, OldFont
    DeleteObject hFont
    
    GetTextSize = s
    dc.Dispose
End Function

Private Sub DrawSeparator(ByVal hdc As Long, ByVal X1 As Long, ByVal X2 As Long, ByVal Y As Long)
    Dim r As RECT
    r.Left = X1
    r.Top = Y
    r.Right = X2
    r.Bottom = Y
    DrawEdge hdc, r, EDGE_ETCHED, BF_TOP
End Sub

Private Sub DrawRoundRect(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal hPen As Long, ByVal hBrush As Long)
    Dim OldBrush As Long, OldPen As Long
    OldPen = SelectObject(hdc, hPen)
    OldBrush = SelectObject(hdc, hBrush)
    RoundRect hdc, X1, Y1, X2, Y2, X3, Y3
    SelectObject hdc, OldPen
    SelectObject hdc, OldBrush
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Private Sub DrawRect(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hPen As Long, ByVal hBrush As Long)
    Dim OldBrush As Long, OldPen As Long
    OldPen = SelectObject(hdc, hPen)
    OldBrush = SelectObject(hdc, hBrush)
    Rectangle hdc, X1, Y1, X2, Y2
    SelectObject hdc, OldPen
    SelectObject hdc, OldBrush
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Private Function SetAllMenusOwnerDraw(ByVal hMenu As Long)
    Dim mii As MENUITEMINFO
    mii.cbSize = Len(mii)
    mii.fMask = MIIM_FTYPE
    
    Dim i As Integer
    
    For i = 0 To GetMenuItemCount(hMenu)
        GetMenuItemInfo hMenu, i, True, mii
        mii.fType = mii.fType Or MF_OWNERDRAW
        SetMenuItemInfo hMenu, i, True, mii
        If MenuHasSubMenues(GetSubMenu(hMenu, i)) Then
            SetAllMenusOwnerDraw GetSubMenu(hMenu, i)
        End If
    Next
End Function

Private Function RemoveAllMenusOwnerDraw(ByVal hMenu As Long)
    Dim mii As MENUITEMINFO
    Dim ODVal As Long
    mii.cbSize = Len(mii)
    mii.fMask = MIIM_FTYPE
    
    Dim i As Integer
    
    For i = 0 To GetMenuItemCount(hMenu)
        GetMenuItemInfo hMenu, i, True, mii
        
        ODVal = mii.fType And MF_OWNERDRAW
        mii.fType = mii.fType - ODVal
        SetMenuItemInfo hMenu, i, True, mii
        If MenuHasSubMenues(GetSubMenu(hMenu, i)) Then
            RemoveAllMenusOwnerDraw GetSubMenu(hMenu, i)
        End If
    Next
End Function

Private Function MenuProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case Is = WM_MENUSELECT
            ApplyMenuSelectEvent hwnd, wParam, lParam
        Case Is = WM_MEASUREITEM
            ApplyMeasureItemStruct hwnd, wParam, lParam
        Case Is = WM_DRAWITEM
            ApplyDrawItemStruct hwnd, wParam, lParam
        Case Is = WM_CLOSE
            SetWindowLong hwnd, GWL_WNDPROC, PrevWndProc
        Case Else
        
    End Select
    MenuProc = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
End Function

Private Function MenuHasSubMenues(ByVal hMenu As Long) As Boolean
    MenuHasSubMenues = GetMenuItemCount(hMenu) <> -1
End Function

Private Function LowWord(Word As Long)
LowWord = CInt("&H" & Right$(Hex$(Word), 4))
End Function

Private Function HighWord(Word As Long)
HighWord = CInt("&H" & Left$(Hex$(Word), 4))
End Function

Private Sub ApplyMenuSelectEvent(ByVal hwnd As Long, ByVal wParam As Long, ByVal lParam As Long)
        Dim Index As Long
        Dim hMenu As Long
        Dim MenuText As String * 255
        
        Index = LowWord(wParam)
        hMenu = lParam
        
        If (HighWord(wParam) And MF_POPUP) = MF_POPUP Then
            GetMenuString lParam, Index, MenuText, 255, MF_BYPOSITION
        
            If Replace(Replace(MenuText, Chr$(0), ""), Chr$(9), "") = "" Then
                frmMain.Menus_MouseSelect Index, True, "", ""
            Else
                If UBound(Split(MenuText, Chr$(9))) = 0 Then
                    frmMain.Menus_MouseSelect Index, True, Replace(MenuText, Chr$(0), ""), ""
                Else
                    frmMain.Menus_MouseSelect Index, True, Replace(Split(MenuText, Chr$(9))(0), Chr$(0), ""), Replace(Split(MenuText, Chr$(9))(1), Chr$(0), "")
                End If
            End If
        Else
            GetMenuString lParam, Index, MenuText, 255, MF_BYCOMMAND
            If Replace(Replace(MenuText, Chr$(0), ""), Chr$(9), "") = "" Then
                frmMain.Menus_MouseSelect Index, True, "", ""
            Else
                If UBound(Split(MenuText, Chr$(9))) = 0 Then
                    frmMain.Menus_MouseSelect Index, False, Replace(MenuText, Chr$(0), ""), ""
                Else
                    frmMain.Menus_MouseSelect Index, False, Replace(Split(MenuText, Chr$(9))(0), Chr$(0), ""), Replace(Split(MenuText, Chr$(9))(1), Chr$(0), "")
                End If
            End If
        End If
End Sub

Private Sub ApplyMeasureItemStruct(ByVal hwnd As Long, ByVal wParam As Long, ByVal lParam As Long)
    Dim mis As MEASUREITEMSTRUCT
    Dim MenuText As String * 255
    Dim i As Integer
    Dim IsMainMenu As Boolean
    Dim MainMenuCaption As String * 255
    
    
    CopyMemory mis, ByVal lParam, Len(mis)
    
    On Error Resume Next
    
    
    If mis.CtlType = ODT_MENU Then
        
        For i = 0 To GetMenuItemCount(GetMenu(hwnd))
            If GetSubMenu(GetMenu(hwnd), i) = mis.itemID Then
                IsMainMenu = True
                Exit For
            End If
        Next
        
        If IsMainMenu Then
            GetMenuString GetMenu(hwnd), mis.itemID, MainMenuCaption, 255, MF_BYCOMMAND
            mis.itemWidth = GetTextSize(Trim(Replace(Replace(MainMenuCaption, Chr(9), ""), Chr(0), "")), msMenuFont).cx
        ElseIf MenuHasSubMenues(mis.itemID) Then ' has sub menus (when the menu has subs, itemID is pointer to it's handle
            mis.itemHeight = msImageHeight + 6
            mis.itemWidth = msImageWidth + 6 + GetTextSize(MyMenus(mis.itemID - 1).MenuText, msMenuFont).cx
        ElseIf MyMenus(mis.itemID - 1).Status = Separator Then
            mis.itemHeight = 5
            mis.itemWidth = 30
        ElseIf MyMenus(mis.itemID - 1).Status = NormalItem Then
            mis.itemHeight = msImageHeight + 6
            mis.itemWidth = msImageWidth + 6 + GetTextSize(MyMenus(mis.itemID - 1).MenuText, msMenuFont).cx
            mis.itemWidth = mis.itemWidth + GetTextSize(MyMenus(mis.itemID - 1).HotKey & Chr$(9), msMenuFont).cx + msImageWidth + 6
            If Len(MyMenus(mis.itemID - 1).HotKey) <> 0 Then mis.itemWidth = mis.itemWidth + msImageWidth + 6
        End If
    End If
    CopyMemory ByVal lParam, mis, Len(mis)
End Sub

Private Sub ApplyDrawItemStruct(ByVal hwnd As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error Resume Next
    Dim dis As DRAWITEMSTRUCT
    
    CopyMemory dis, ByVal lParam, Len(dis)
    
    If dis.CtlType = ODT_MENU Then
        If GetMenu(hwnd) = dis.hwndItem Then ' is in menu bar
            DrawMainMenuItem hwnd, dis
            
        ElseIf MenuHasSubMenues(dis.itemID) Then ' has sub menus (when the menu has subs, itemID is pointer to it's handle
            DrawItem dis
        ElseIf MyMenus(dis.itemID - 1).Status = Separator Then
            DrawSeparatorItem dis
        ElseIf MyMenus(dis.itemID - 1).Status = NormalItem Then
            DrawItem dis
        End If
    End If
    
    'dont let the windows draw onther things on menu, such as the triangle pointer to a sub menu
    If dis.CtlType = ODT_MENU Then _
    ExcludeClipRect dis.hdc, dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Right, dis.rcItem.Bottom
    
    CopyMemory ByVal lParam, dis, Len(dis)
End Sub

Private Sub DrawItem(ByRef dis As DRAWITEMSTRUCT)
    If dis.itemState And ODS_CHECKED Then
        If dis.itemState And ODS_SELECTED Then
            DrawSelectedState dis
        Else
            DrawNormalState dis
        End If
            DrawCheckIcon dis
    ElseIf dis.itemState And ODS_DISABLED Then
        DrawDisabledState dis
        DrawImageIcon dis
    ElseIf dis.itemState And ODS_SELECTED Then
        DrawSelectedState dis
        DrawImageIcon dis
    Else
        DrawNormalState dis
        DrawImageIcon dis
    End If
End Sub

Private Sub DrawSeparatorItem(ByRef dis As DRAWITEMSTRUCT)
    Dim g As New Gradient
    Dim dc As New DeviceContext
    Dim r As RECT
    Dim i As Integer
    Dim HasSubMenus As Boolean
    Dim hPen As Long, hBrush As Long
    Dim dt As New DrawText
    
    Dim pic As IPictureDisp
    r = dis.rcItem
    
    hPen = CreatePen(PS_NULL, 0, 0)
    hBrush = CreateSolidBrush(vbWhite)
    DrawRect dis.hdc, r.Left, r.Top, r.Right + 1, r.Bottom, hPen, hBrush
    g.Rectangle dis.hdc, r.Left, r.Top - 1, msImageWidth + 6, r.Bottom + 2, RGB(215, 233, 255), RGB(111, 164, 200), False
    g.Rectangle dis.hdc, r.Left, r.Top - 1, msImageWidth + 6, r.Bottom + 2, RGB(215, 233, 255), RGB(111, 164, 200), False
    DrawSeparator dis.hdc, msImageWidth + 6 + 3, r.Right - 3, (r.Top + r.Bottom) / 2
End Sub

Private Sub DrawSelectedState(ByRef dis As DRAWITEMSTRUCT)
    Dim g As New Gradient
    Dim dc As New DeviceContext
    Dim r As RECT
    Dim i As Integer
    Dim HasSubMenus As Boolean
    Dim hPen As Long, hBrush As Long
    Dim dt As New DrawText
    
    Dim pic As IPictureDisp
    Dim MenuCaption As String
    
    r = dis.rcItem
    
    On Error Resume Next
    

    If IsMenuHandle(dis.itemID) Then
        MenuCaption = Space(255)
        GetMenuString dis.hwndItem, dis.itemID, MenuCaption, 255, MF_BYCOMMAND
        MenuCaption = Trim(Replace(MenuCaption, Chr(0), ""))
    Else
        MenuCaption = MyMenus(dis.itemID - 1).MenuText
    End If
    
    dc.Create 24, r.Right, r.Bottom
    g.Rectangle dc.Handle, r.Left, r.Top, r.Right, r.Bottom, RGB(255, 255, 128), RGB(255, 128, 64), True
    hPen = CreatePen(0, 1, RGB(255, 192, 192))
    hBrush = CreatePatternBrush(dc.ConvertToBitmap(0, 0, dc.Width, dc.Height))
    DrawRoundRect dis.hdc, r.Left, r.Top, r.Right, r.Bottom, 2, 2, hPen, hBrush
    dc.Dispose
    
    dt.Align = vbLeftJustify
    dt.hdc = dis.hdc
    dt.Prefix = ShowPrefix
    dt.SingleLine = True
    dt.VerticalAlign = TextVerticalAlign.AlignVCenter
    dt.SmoothingMode = SmoothingModeClearType
    
    dt.Draw Split(MenuCaption, Chr(9))(0), r.Left + msImageWidth + 6 + 3, r.Top, r.Right, r.Bottom, msMenuFont, vbBlack, 0, True, False
    
    dt.Align = vbRightJustify
    dt.Draw MyMenus(dis.itemID - 1).HotKey, r.Left + msImageWidth + 6 + 3, r.Top, r.Right - msImageWidth - 6, r.Bottom, msMenuFont, vbBlack, 0, True, False
    
    If IsMenuHandle(dis.itemID) Then g.Triangle dis.hdc, r.Right - 10, r.Top + 6, RGB(96, 96, 96), _
                                                                          r.Right - 9, r.Top + 14, RGB(96, 96, 96), _
                                                                          r.Right - 5, r.Top + 10, RGB(96, 96, 96)
                                                                          
End Sub

Private Sub DrawNormalState(ByRef dis As DRAWITEMSTRUCT)
    Dim g As New Gradient
    Dim dc As New DeviceContext
    Dim r As RECT
    Dim i As Integer
    Dim HasSubMenus As Boolean
    Dim hPen As Long, hBrush As Long
    Dim dt As New DrawText
    Dim points(0 To 2) As POINTF
    Dim pic As IPictureDisp
    Dim MenuCaption As String
    
    r = dis.rcItem
    
    On Error Resume Next
    
    
    If IsMenuHandle(dis.itemID) Then
        MenuCaption = Space(255)
        GetMenuString dis.hwndItem, dis.itemID, MenuCaption, 255, MF_BYCOMMAND
        MenuCaption = Trim(Replace(MenuCaption, Chr(0), ""))
    
    Else
        MenuCaption = MyMenus(dis.itemID - 1).MenuText
    End If
        
    hBrush = CreateSolidBrush(vbWhite)
    hPen = CreatePen(PS_NULL, 1, vbWhite)
    DrawRect dis.hdc, r.Left - 1, r.Top - 1, r.Right + 2, r.Bottom + 2, hPen, hBrush
    
    dt.Align = vbLeftJustify
    dt.hdc = dis.hdc
    dt.Prefix = ShowPrefix
    dt.SingleLine = True
    dt.VerticalAlign = TextVerticalAlign.AlignVCenter
    dt.SmoothingMode = SmoothingModeClearType

    dt.Draw Split(MenuCaption, Chr(9))(0), r.Left + msImageWidth + 6 + 3, r.Top, r.Right, r.Bottom, msMenuFont, vbBlack, 0, True, False
    
    dt.Align = vbRightJustify
    g.Rectangle dis.hdc, r.Left, r.Top - 1, msImageWidth + 6, r.Bottom + 2, RGB(215, 233, 255), RGB(111, 164, 200), False
    dt.Draw MyMenus(dis.itemID - 1).HotKey, r.Left + msImageWidth + 6 + 3, r.Top, r.Right - msImageWidth - 6, r.Bottom, msMenuFont, vbBlack, 0, True, False


    If IsMenuHandle(dis.itemID) Then g.Triangle dis.hdc, r.Right - 10, r.Top + 6, RGB(96, 96, 96), _
                                                                          r.Right - 9, r.Top + 14, RGB(96, 96, 96), _
                                                                          r.Right - 5, r.Top + 10, RGB(96, 96, 96)
End Sub

Private Sub DrawDisabledState(ByRef dis As DRAWITEMSTRUCT)
    Dim g As New Gradient
    Dim dc As New DeviceContext
    Dim r As RECT
    Dim i As Integer
    Dim HasSubMenus As Boolean
    Dim hPen As Long, hBrush As Long
    Dim dt As New DrawText
    
    Dim pic As IPictureDisp
    Dim MenuCaption As String
    
    r = dis.rcItem
    
    On Error Resume Next
    
    If IsMenuHandle(dis.itemID) Then
        MenuCaption = Space(255)
        GetMenuString dis.hwndItem, dis.itemID, MenuCaption, 255, MF_BYCOMMAND
        MenuCaption = Trim(Replace(MenuCaption, Chr(0), ""))
    Else
        MenuCaption = MyMenus(dis.itemID - 1).MenuText
    End If
    
    hBrush = CreateSolidBrush(vbWhite)
    hPen = CreatePen(PS_NULL, 1, vbWhite)
    DrawRect dis.hdc, r.Left - 1, r.Top - 1, r.Right + 2, r.Bottom + 2, hPen, hBrush
    
    dt.Align = vbLeftJustify
    dt.hdc = dis.hdc
    dt.Prefix = ShowPrefix
    dt.SingleLine = True
    dt.VerticalAlign = TextVerticalAlign.AlignVCenter
    dt.SmoothingMode = SmoothingModeClearType

    dt.Draw Split(MenuCaption, Chr(9))(0), r.Left + msImageWidth + 6 + 3, r.Top, r.Right, r.Bottom, msMenuFont, RGB(128, 128, 128), 0, True, False
    
    dt.Align = vbRightJustify
    g.Rectangle dis.hdc, r.Left, r.Top - 1, msImageWidth + 6, r.Bottom + 2, RGB(215, 233, 255), RGB(111, 164, 200), False
    dt.Draw MyMenus(dis.itemID - 1).HotKey, r.Left + msImageWidth + 6 + 3, r.Top, r.Right - msImageWidth - 6, r.Bottom, msMenuFont, RGB(128, 128, 128), 0, True, False
End Sub

Private Sub DrawImageIcon(ByRef dis As DRAWITEMSTRUCT)
    
    Dim g As New Gradient
    Dim dc As New DeviceContext
    Dim hPen As Long, hBrush As Long
    Dim r As RECT
    
    Dim MenuHandle As Long
    Dim ImgIndex As Long
    Dim i As Long

    
    r = dis.rcItem
    
    If IsSubMenuHandle(dis.hwndItem, dis.itemID) Then
    
    
        For i = 0 To UBound(MyMenus)
            If MyMenus(i).MenuHandle = dis.itemID Then
                ImgIndex = MyMenus(i).MenuIndex
                Exit For
            End If
        Next
    Else
        ImgIndex = MyMenus(dis.itemID).MenuIndex
    End If
    
    On Error Resume Next
    If Len(msImageFiles(ImgIndex)) <> 0 Then
        If Len(ImgIndex) <> 0 Then
            gp.LoadPictureToHDC msImageFiles(ImgIndex), dis.hdc, r.Left + 3, r.Top + 3, msImageWidth, msImageHeight, 0, 0, msImageWidth, msImageHeight
        End If
    End If
End Sub

Private Sub DrawImageIconDisabled(ByRef dis As DRAWITEMSTRUCT)

End Sub

Private Sub DrawImageCheckIcon(ByRef dis As DRAWITEMSTRUCT)
    Dim g As New Gradient
    Dim dc As New DeviceContext
    Dim r As RECT
    Dim hPen As Long, hBrush As Long
    r = dis.rcItem
    
    On Error Resume Next
    
    If MyMenus(dis.itemID).Status = NormalItem Then
    End If
End Sub

Public Function GetMenuCount() As Long
    Dim mnu() As MenuType
    GetMenuStrings GetMenu(mshWnd), mnu
    GetMenuCount = UBound(mnu)
End Function

Private Sub DrawCheckIcon(ByRef dis As DRAWITEMSTRUCT)
    Dim hPen As Long, hBrush As Long
    Dim dc As New DeviceContext
    Dim g As New Gradient
    
    Dim dt As New DrawText, TickFont As New StdFont
    Dim r As RECT
    
    r = dis.rcItem
    
    If Len(msImageFiles(dis.itemID)) <> 0 Then
        
        dc.Create 24, msImageWidth + 6, msImageHeight + 6
        g.Rectangle dc.Handle, 0, 0, msImageWidth + 6, msImageHeight + 6, RGB(222, 184, 135), RGB(245, 245, 220), True
        hPen = CreatePen(0, 1, RGB(120, 140, 160))
        hBrush = CreatePatternBrush(dc.ConvertToBitmap(0, 0, dc.Width, dc.Height))
        DrawRoundRect dis.hdc, r.Left + 2, r.Top + 2, r.Left + msImageWidth + 6 - 2, r.Top + msImageHeight + 6 - 2, 4, 4, hPen, hBrush
        dc.Dispose
        
        gp.LoadPictureToHDC msImageFiles(dis.itemID), dis.hdc, r.Left + 3, r.Top + 3, msImageWidth, msImageHeight, 0, 0, msImageWidth, msImageHeight
        
    Else
        dc.Create 24, msImageWidth + 6, msImageHeight + 6
        
        g.Rectangle dc.Handle, 0, 0, dc.Width, dc.Height, RGB(222, 184, 135), RGB(245, 245, 220), True
        
        hPen = CreatePen(0, 1, RGB(245, 160, 45))
        hBrush = CreatePatternBrush(dc.ConvertToBitmap(3, 3, dc.Width - 3, dc.Height - 3))
        
        dc.GetFromDC dis.hdc, r.Left, r.Top, r.Right, r.Bottom, vbSrcCopy
        DrawRoundRect dc.Handle, 3, 3, dc.Width - 3, dc.Height - 3, 2, 2, hPen, hBrush
        
        dt.Align = vbCenter
        dt.SingleLine = True
        dt.VerticalAlign = AlignVCenter
        dt.MeasureInPexils = True
        dt.hdc = dc.Handle
        dt.SmoothingMode = SmoothingModeClearType
        
        With TickFont
            .Name = "Tahoma"
            .Size = dc.Height
        End With
        
        dt.Draw ChrW(&H2713), 0, 0, dc.Width, dc.Height, TickFont, vbBlack, 0, True, True
        dc.SetToDC dis.hdc, r.Left, r.Top, msImageWidth + 6, msImageHeight + 6, 0, 0, dc.Width, dc.Height, Qualities.QualityHalftoneOrBilinear, vbSrcCopy
        
        dc.Dispose
    End If
End Sub

Private Sub DrawMainMenuItem(ByVal hwnd As Long, ByRef dis As DRAWITEMSTRUCT)
    Dim ist As Long
    ist = dis.itemState
    
    If ist And ODS_INACTIVE Then 'the owner window is not active
        If ist And ODS_HOTLIGHT Then
            DrawHotLightMainMenuItem hwnd, dis
        ElseIf ist And ODS_DISABLED Then
            DrawDisabledMainMenuItem hwnd, dis
        Else
            DrawInactiveMainMenuItem hwnd, dis
        End If
    Else
        If ist And ODS_HOTLIGHT Then ' mouse hover
            DrawHotLightMainMenuItem hwnd, dis
        ElseIf ist And ODS_DISABLED Then
            DrawDisabledMainMenuItem hwnd, dis
        ElseIf ist And ODS_SELECTED Then 'selected by clicking on it
            DrawSelectedMainMenuItem hwnd, dis
        Else ' normal mode
            DrawNormalMainMenuItem hwnd, dis
        End If
    End If
End Sub

Private Sub DrawNormalMainMenuItem(ByVal hwnd As Long, ByRef dis As DRAWITEMSTRUCT)
    DrawGraphicalMainMenuItem hwnd, dis, RGB(215, 233, 255), RGB(215, 233, 255), False, 0, 0, _
    0, 0, 0, 0, vbBlack
End Sub

Private Sub DrawInactiveMainMenuItem(ByVal hwnd As Long, ByRef dis As DRAWITEMSTRUCT)
    DrawGraphicalMainMenuItem hwnd, dis, RGB(215, 233, 255), RGB(215, 233, 255), False, 0, 0, _
    0, 0, 0, 0, RGB(96, 96, 96)
End Sub

Private Sub DrawDisabledMainMenuItem(ByVal hwnd As Long, ByRef dis As DRAWITEMSTRUCT)
    DrawGraphicalMainMenuItem hwnd, dis, RGB(215, 233, 255), RGB(215, 233, 255), False, 0, 0, _
    0, 0, 0, 0, RGB(128, 128, 128)
End Sub

Private Sub DrawHotLightMainMenuItem(ByVal hwnd As Long, ByRef dis As DRAWITEMSTRUCT)
    DrawGraphicalMainMenuItem hwnd, dis, RGB(215, 233, 255), RGB(215, 233, 255), True, 2, RGB(255, 128, 0), _
    1, 1, RGB(255, 180, 60), RGB(255, 130, 40), vbBlack
End Sub

Private Sub DrawSelectedMainMenuItem(ByVal hwnd As Long, ByRef dis As DRAWITEMSTRUCT)
    DrawGraphicalMainMenuItem hwnd, dis, RGB(215, 233, 255), RGB(215, 233, 255), True, 2, RGB(255, 128, 0), _
    1, 1, RGB(255, 180, 40), RGB(240, 180, 120), vbBlack
End Sub

Private Sub DrawGraphicalMainMenuItem(ByVal hwnd As Long, ByRef dis As DRAWITEMSTRUCT, ByVal GradientColor1 As Long, ByVal GradientColor2 As Long, ByVal RoundedRect As Boolean, ByVal RoundRectRadius As Long, ByVal RoundRectBorderColor As Long, ByVal RoundRectBorderWidth As Long, ByVal RoundRectIndent As Long, ByVal RoundRectColor1 As Long, ByVal RoundRectColor2 As Long, ByVal TextColor As Long)
    Dim g As New Gradient
    Dim dc As New DeviceContext
    Dim r As RECT
    Dim i As Integer
    Dim HasSubMenus As Boolean
    Dim hPen As Long, hBrush As Long
    Dim dt As New DrawText
    
    Dim pic As IPictureDisp
    Dim MenuCaption As String * 255
    Dim MenuCaptionCorrect As String
    Dim MenuSize As Size
    On Error Resume Next                                                                                                                                                                                                                                                                                                                                                                           ' _
    when the menu item is in main menu, the .hwndItem is handle to parent menu, and .itemID is the                                                                                                                                                                                                                                                                                                                                           ' _
    handle to menu item
    GetMenuString dis.hwndItem, dis.itemID, MenuCaption, 255, MF_BYCOMMAND
    MenuCaptionCorrect = Trim(Replace(Replace(MenuCaption, Chr(9), ""), Chr(0), ""))
    MenuSize = GetTextSize(MenuCaptionCorrect, msMenuFont)
    
    r = dis.rcItem
    Dim rct As RECT
    If dis.itemID = GetSubMenu(GetMenu(hwnd), GetMenuItemCount(GetMenu(hwnd)) - 1) Then _
    GetMenuItemRect hwnd, GetMenu(hwnd), 0, rct
    
    g.Rectangle dis.hdc, r.Left, r.Top, r.Right, r.Bottom, GradientColor1, GradientColor2, True
        
    If RoundedRect Then
        dc.Create 24, r.Right, r.Bottom
        g.Rectangle dc.Handle, r.Left, r.Top, dc.Width, dc.Height, RoundRectColor1, RoundRectColor2, True
        If RoundRectBorderWidth > 0 Then
            hPen = CreatePen(0, RoundRectBorderWidth, RoundRectBorderColor)
        Else
            hPen = CreatePen(PS_NULL, RoundRectBorderWidth, RoundRectBorderColor)
        End If
        hBrush = CreatePatternBrush(dc.ConvertToBitmap(0, 0, dc.Width, dc.Height))
        DrawRoundRect dis.hdc, r.Left + RoundRectIndent, r.Top + RoundRectIndent, r.Right - RoundRectIndent, r.Bottom - RoundRectIndent, RoundRectRadius, RoundRectRadius, hPen, hBrush
        dc.Dispose
    End If
    
    dt.hdc = dis.hdc
    dt.Align = vbCenter
    dt.Prefix = ShowPrefix
    dt.SingleLine = True
    dt.VerticalAlign = TextVerticalAlign.AlignVCenter
    dt.SmoothingMode = SmoothingModeAntiAliased

    dt.Draw MenuCaptionCorrect, r.Left, r.Top, r.Right, r.Bottom, msMenuFont, TextColor, 0, True, False
    Exit Sub
err1:
    MsgBox Err.description
End Sub

Private Function GetMenuStrings(hMenu As Long, RetVal() As MenuType)

    Dim i As Long
    Static cnt As Long
    Static MenuItemText As String
    Dim CurrentMenuStatus As MenuStatus
    
    Static MenuIndexes As String
    Static MenuTexts As String
    Static MenuHandles As String
    Static MenuHotKeys As String
    Static MenusStatus As String
    
    Static MenuIndexesArrayStr() As String
    Static MenuHandlesArrayStr() As String
    Static MenuTextsArrayStr() As String
    Static MenuHotKeysArrayStr() As String
    Static MenusStatusArrayStr() As String
    
    Static MenuIndexesArray() As Long
    Static MenuHandlesArray() As Long
    Static MenuTextsArray() As String
    Static MenuHotKeysArray() As String
    Static MenusStatusArray() As Long
    
    For i = 0 To GetMenuItemCount(hMenu) - 1
        
        cnt = cnt + 1
        MenuItemText = Space$(255)
        GetMenuString hMenu, i, MenuItemText, 255, MF_BYPOSITION
        MenuItemText = Trim(Replace(MenuItemText, Chr(0), ""))
        
        
        If MenuHasSubMenues(GetSubMenu(hMenu, i)) Then
            CurrentMenuStatus = HasSubs
        ElseIf IsMenuSeparator(hMenu, i, True) Then
            CurrentMenuStatus = Separator
        Else
            CurrentMenuStatus = NormalItem
        End If
        
        MenusStatus = MenusStatus & CurrentMenuStatus & vbNewLine
        
        If MenuItemText = "" Then
            MenuTexts = MenuTexts & "" & vbNewLine
            MenuHotKeys = MenuHotKeys & "" & vbNewLine
        Else
            If UBound(Split(MenuItemText, Chr(9))) = 0 Then
                MenuTexts = MenuTexts & MenuItemText & vbNewLine
                MenuHotKeys = MenuHotKeys & "" & vbNewLine
            Else
                MenuTexts = MenuTexts & Split(MenuItemText, Chr(9))(0) & vbNewLine
                MenuHotKeys = MenuHotKeys & Split(MenuItemText, Chr(9))(1) & vbNewLine
            End If
        End If
        
        MenuIndexes = MenuIndexes & cnt & vbNewLine
        MenuHandles = MenuHandles & GetSubMenu(hMenu, i) & vbNewLine
        
        If MenuHasSubMenues(GetSubMenu(hMenu, i)) Then
            GetMenuStrings GetSubMenu(hMenu, i), RetVal
        End If
    Next
    
    MenuIndexesArrayStr = Split(MenuIndexes, vbNewLine)
    MenuHandlesArrayStr = Split(MenuHandles, vbNewLine)
    MenuTextsArrayStr = Split(MenuTexts, vbNewLine)
    MenuHotKeysArrayStr = Split(MenuHotKeys, vbNewLine)
    MenusStatusArrayStr = Split(MenusStatus, vbNewLine)
    
    ReDim Preserve MenuIndexesArray(UBound(MenuIndexesArrayStr))
    ReDim Preserve MenuHandlesArray(UBound(MenuHandlesArrayStr))
    ReDim Preserve MenuTextsArray(UBound(MenuTextsArrayStr))
    ReDim Preserve MenuHotKeysArray(UBound(MenuHotKeysArrayStr))
    ReDim Preserve MenusStatusArray(UBound(MenusStatusArrayStr))
    
    For i = 0 To UBound(MenuIndexesArrayStr)
        MenuIndexesArray(i) = Val(MenuIndexesArrayStr(i))
        MenuHandlesArray(i) = Val(MenuHandlesArrayStr(i))
        MenuTextsArray(i) = (MenuTextsArrayStr(i))
        MenuHotKeysArray(i) = (MenuHotKeysArrayStr(i))
        MenusStatusArray(i) = Val(MenusStatusArrayStr(i))
    Next
    
    ReDim Preserve RetVal(UBound(Split(MenuTexts, vbNewLine)))
    For i = 0 To UBound(Split(MenuTexts, vbNewLine)) - 1
        RetVal(i).MenuIndex = MenuIndexesArray(i)
        RetVal(i).MenuHandle = MenuIndexesArray(i)
        RetVal(i).MenuText = MenuTextsArray(i)
        RetVal(i).HotKey = MenuHotKeysArray(i)
        RetVal(i).Status = MenusStatusArray(i)
    Next
    ReDim Preserve RetVal(UBound(RetVal) - 1)
End Function

Private Function IsMenuSeparator(ByVal hMenu As Long, ByVal SubMenu As Long, ByVal ByPoision As Boolean) As Boolean
    Dim mii As MENUITEMINFO
    mii.cbSize = Len(mii)
    mii.fMask = MIIM_FTYPE
     GetMenuItemInfo hMenu, SubMenu, ByPoision, mii
    IsMenuSeparator = ((mii.fType And MF_SEPARATOR) = MF_SEPARATOR)
End Function

Private Function IsSubMenuHandle(ByVal hMenu As Long, ByVal hSubMenu As Long) As Boolean
    If Not (GetMenuState(hMenu, hSubMenu, MF_BYCOMMAND) And MF_POPUP) Then
        IsSubMenuHandle = True
    Else
        IsSubMenuHandle = False
    End If
End Function

Private Function IsMenuHandle(ByVal hMenu As Long) As Boolean
    If GetMenuItemCount(hMenu) = -1 Then
        IsMenuHandle = False
    Else
        IsMenuHandle = True
    End If
End Function

