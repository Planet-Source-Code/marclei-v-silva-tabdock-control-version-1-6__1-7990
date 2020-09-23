Attribute VB_Name = "modApi32"
' ******************************************************************************
' Module      : modApi32.bas
' Created by  : Marclei V Silva
' Machine     : ZEUS
' Date-Time   : 09/05/20003:09:33
' Description : Several Api declares, constants and definitions
' ******************************************************************************
Option Explicit

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

' System metrics constants
Public Const SM_CXSIZE = 30
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CYSIZE = 31
Public Const SM_CYCAPTION = 4
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CYMENU = 15
Public Const SM_CYSMCAPTION = 51 'height of windows 95 small caption

' These constants define the style of border to draw.
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2

Public Const BF_FLAT = &H4000
Public Const BF_MONO = &H8000
Public Const BF_SOFT = &H1000      ' For softer buttons

Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

' These constants define which sides to draw.
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const SWP_NOOWNERZORDER = &H200              ' Don"t do owner Z ordering
Public Const SWP_FRAMECHANGED = &H20                ' The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HWNDPARENT = (-8)

Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1

' Window styles
Public Const WS_ACTIVECAPTION = &H1
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000 'WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_GT = WS_GROUP Or WS_TABSTOP
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000

Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_RBUTTONUP = &H205

' Extended window styles
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_CONTEXTHELP = &H400&
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_LAYOUTRTL = &H400000 ' Right to left mirroring
Public Const WS_EX_LEFT = &H0&
Public Const WS_EX_LEFTSCROLLBAR = &H4000&
Public Const WS_EX_LTRREADING = &H0&
Public Const WS_EX_MDICHILD = &H40&
Public Const WS_EX_NOACTIVATE = &H8000000
Public Const WS_EX_NOINHERITLAYOUT = &H100000 ' Disable inheritence of mirroring by children
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_RIGHT = &H1000&
Public Const WS_EX_RIGHTSCROLLBAR = &H0&
Public Const WS_EX_RTLREADING = &H2000&
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_TOOLWINDOW = &H80&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
Public Const WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST

Public Const SC_CLOSE = &HF060&
Public Const SC_MOVE = &HF010&
Public Const SC_SIZE = &HF000&

Public Const OPAQUE = 2
Public Const VK_LBUTTON = &H1
Public Const PS_SOLID = 0
Public Const BLACK_PEN = 7
Public Const MOUSE_MOVE = &HF012
Public Const TRANSPARENT = 1
Public Const BITSPIXEL = 12

' subclassing constants
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_MOVE = &H3
Public Const WM_EXITSIZEMOVE = &H232
Public Const WM_SIZE = &H5
Public Const WM_USER = &H400
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_SYSCOMMAND = &H112
Public Const WM_NULL = &H0
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_ACTIVATE As Long = &H6
Public Const WM_KILLFOCUS = &H8
Public Const WM_PAINT = &HF
Public Const WM_DESTROY = &H2
Public Const WM_NCHITTEST = &H84
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_COMMAND = &H111
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_STYLECHANGED As Long = &H7D&
'Public Const WM_DESTROY = &H2
Public Const WM_SIZING = &H214
Public Const WM_MOVING = &H216&
Public Const WM_ENTERSIZEMOVE = &H231&
'Public Const WM_EXITSIZEMOVE = &H232&
'Public Const WM_ACTIVATE = &H6
'Public Const WM_SIZE = &H5
Public Const WM_CLOSE = &H10

Public Const PS_INSIDEFRAME = 6

' Region constants
Public Const RGN_OR = 2     ' RGN_OR creates the union of combined regions
Public Const RGN_DIFF = 4   ' RGN_DIFF creates the intersection of combined regions
Public Const RGN_AND = 1
Public Const RGN_XOR = 3

' SysCommand
Public Const HTCAPTION = 2
Public Const HTCLOSE = 20

Public Const R2_BLACK = 1       '   0
Public Const R2_COPYPEN = 13    '  P
Public Const R2_LAST = 16
Public Const R2_MASKNOTPEN = 3  '  DPna
Public Const R2_MASKPEN = 9     '  DPa
Public Const R2_MASKPENNOT = 5  '  PDna
Public Const R2_MERGENOTPEN = 12        '  DPno
Public Const R2_MERGEPEN = 15   '  DPo
Public Const R2_MERGEPENNOT = 14        '  PDno
Public Const R2_NOP = 11        '  D
Public Const R2_NOT = 6 '  Dn
Public Const R2_NOTCOPYPEN = 4  '  PN
Public Const R2_NOTMASKPEN = 8  '  DPan
Public Const R2_NOTMERGEPEN = 2 '  DPon
Public Const R2_NOTXORPEN = 10  '  DPxn
Public Const R2_WHITE = 16      '   1
Public Const R2_XORPEN = 7      '  DPx

Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Public Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal lLeft As Long, ByVal lTop As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' ******************************************************************************
' Routine       : ObjectFromPtr
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:17:24
' Inputs        : lPtr - pointer to the object
' Outputs       : An object
' Credits       : SP MacMahon (www.vbaccelerator.com articles)
' Modifications : None
' Description   : Get an object from the given pointer
' ******************************************************************************
Public Function ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oThis As Object

    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory oThis, lPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oThis
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oThis, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but this will be your code rather than the uncounted reference!

End Function

' ******************************************************************************
' Routine       : PtrFromObject
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:19:00
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Get a pointer fro a object
' ******************************************************************************
Public Function PtrFromObject(ByRef oThis As Object) As Long
    ' Return the pointer to this object:
    PtrFromObject = ObjPtr(oThis)
End Function

Function HiWord(ByVal dw As Long) As Integer
   If dw And &H80000000 Then
         HiWord = (dw \ 65535) - 1
   Else: HiWord = dw \ 65535
   End If
End Function

Function LoWord(ByVal dw As Long) As Integer
   If dw And &H8000& Then
         LoWord = &H8000 Or (dw And &H7FFF&)
   Else: LoWord = dw And &HFFFF&
   End If
End Function
'-- end code
