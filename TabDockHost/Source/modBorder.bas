Attribute VB_Name = "modBorder"
'
'  Color your Border
'  Code originally intended to produce 3D borders
'       for WIN 3.x Forms.
'  Reapplied by: linda
'  linda.69@mailcity.com
'
'  can be easily modified to draw anything on the
'  title bar of the form.
'

Option Explicit
DefInt A-Z

' private declares
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long) As Long

' Pen Styles
Private Const PS_SOLID = 0
Private Const CLR_INVALID = 0

' ******************************************************************************
' Routine       : DrawBorder
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 02/10/0010:20:55
' Inputs        :
' Outputs       :
' Credits       : linda.69@mailcity.com (Color your Border demo)
' Modifications : Color translation OLE_COLOR to RGB
' Description   : draw a user defined color border
' ******************************************************************************
Public Sub DrawBorder(frmTarget As Form, Color As OLE_COLOR)
    Dim hWindowDC As Long
    Dim hOldPen As Long
    Dim nLeft As Long
    Dim nRight As Long
    Dim nTop As Long
    Dim nBottom As Long
    Dim Ret As Long
    Dim hMyPen As Long
    Dim WidthX As Long
    Dim rgbColor As Long
    
    ' translate
    rgbColor = TranslateColor(Color)
    ' border width
    WidthX = GetSystemMetrics(SM_CYBORDER) * 5
    ' get window DC
    hWindowDC = GetWindowDC(frmTarget.hWnd)   'this is outside the form
    ' create a pen
    hMyPen = CreatePen(PS_SOLID, WidthX, rgbColor)
    ' Initialize misc variables
    nLeft = 0: nTop = 0
    nRight = frmTarget.Width / Screen.TwipsPerPixelX
    nBottom = frmTarget.Height / Screen.TwipsPerPixelY
    ' select border pen
    hOldPen = SelectObject(hWindowDC, hMyPen)
    ' draw color around the border
    Ret = LineTo(hWindowDC, nLeft, nBottom)
    Ret = LineTo(hWindowDC, nRight, nBottom)
    Ret = LineTo(hWindowDC, nRight, nTop)
    Ret = LineTo(hWindowDC, nLeft, nTop)
    ' select old pen
    Ret = SelectObject(hWindowDC, hOldPen)
    Ret = DeleteObject(hMyPen)
    Ret = ReleaseDC(frmTarget.hWnd, hWindowDC)
End Sub

' ******************************************************************************
' Routine       : TranslateColor
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 02/10/0010:20:19
' Inputs        :
' Outputs       :
' Credits       : Extracted from VB KB Article
' Modifications :
' Description   : Converts an OLE_COLOR to RGB color
' ******************************************************************************
Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hpal As Long = 0) As Long
    If OleTranslateColor(clr, hpal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function
