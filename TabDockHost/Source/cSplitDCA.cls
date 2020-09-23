VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSplitDDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ======================================================================
' Class    : cSplitDDC
' Filename : cSplitDC.cls
' Author   : SP McMahon
' Date     : 07 July 1998
'
' A splitter class using the Desktop window to draw a
' splitter bar, therefore allowing splitting of MDI forms
' as well as standard forms.
' ======================================================================

' REMARKS
' This code was extract from vbaccelarator web site and modified
' to acomplish the requirements of this project
'
' Thanks! to SP McMahon
'
' Modifications
' -------------
'
' -  The splitter object (and further calls to it) was replaced
'    by Top and Left properties once the Userobject TabDockHost
'    is already the bounds for splitting actions
'
' -  The events FormMouseMove and FormMouseUp was put together
'    in the FormMouseDown Event to supply a immediate split
'    drawing for TabDock UserControl
'
' -  The Clipping area (MouseDown() event) is now handled by
'    the TabDockHost MouseDown() event once this event calculates
'    the rectangle of the docked window
'
' -  Removed Cursor clipping (1.6)
'
' -  Removed Border references. This project does not need (1.6)
'
' -  Window Rect was replace for a Rect passed as reference
'    which is the exact area we want to split (1.6)
'
' -  Created an offset based on start and end cursor position
' removed declarations to the modAPi32 for reducing project size
'*******************
Option Explicit

'// some global declarations
Private bDraw As Boolean
Private rcCurrent As RECT
Private rcNew As RECT
Private rcWindow As RECT
Private m_Offset As Long

Public Enum eOrientationConstants
    espVertical = 1
    espHorizontal = 2
End Enum

Private m_hWnd As Long
Private m_eOrientation As eOrientationConstants

Private m_bSplitting As Boolean

Public Property Get Offset() As Long
    Offset = m_Offset
End Property

Public Property Get Orientation() As eOrientationConstants
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal eOrientation As eOrientationConstants)
    m_eOrientation = eOrientation
End Property

Friend Sub SplitterMouseDown(ByVal hWnd As Long, Rc As RECT, ByVal x As Long, ByVal Y As Long)
    Dim tP As POINTAPI
    Dim tpPrev As POINTAPI
    Dim hDC As Long
    Dim hPen As Long
    
    m_hWnd = hWnd
    ' Send subsequent mouse messages to the owner window
    SetCapture m_hWnd
    ' get window rect
    rcWindow = Rc
    ' Get the Desktop DC:
'    hDC = CreateDCAsNull("DISPLAY", 0, 0, 0)
    ' Set to XOR drawing mode:
'    SetROP2 hDC, R2_NOTXORPEN
    'hPen = CreatePen(2, 1, vbButtonFace)
    'SelectObject hDC, hPen
    ' Erase the last rectangle:
'    Rectangle hDC, Rc.Left, Rc.Top, Rc.Right, Rc.Bottom
    ' Clear up the desktop DC:
'    DeleteDC hDC
    'DeleteObject hPen
    GetCursorPos tP
    ' Store the initial cursor position
    tpPrev.x = tP.x
    tpPrev.Y = tP.Y
    If (m_eOrientation = espHorizontal) Then
        m_Offset = tP.Y
    Else
        m_Offset = tP.x
    End If
    bDraw = True  ' start actual drawing from next move message
    rcCurrent.Left = 0: rcCurrent.Top = 0: rcCurrent.Right = 0: rcCurrent.Bottom = 0
    SplitterFormMouseMove tP.x, tP.Y
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos tP
        If tP.x <> tpPrev.x Or tP.Y <> tpPrev.Y Then
            tpPrev.x = tP.x
            tpPrev.Y = tP.Y
            SplitterFormMouseMove tP.x, tP.Y
        End If
        DoEvents
    Loop
    SplitterFormMouseUp tP.x, tP.Y
End Sub

' Changed to private
Private Sub SplitterFormMouseMove(ByVal x As Long, ByVal Y As Long)
    Dim hDC As Long
    Dim tP As POINTAPI
    'Dim hPen As Long
    
    If (bDraw) Then
        ' Draw two rectangles in the screen DC to cause splitting:
        ' First get the Desktop DC:
        hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        ' Set the draw mode to XOR:
        SetROP2 hDC, R2_NOTXORPEN
        'hPen = CreatePen(2, 2, vbButtonFace)
        'SelectObject hDC, hPen
        ' Draw over and erase the old rectangle
        ' (if this is the first time, all the coords will be 0 and nothing will get drawn):
        Rectangle hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
        ' It is simpler to use the mouse cursor position than try to translate
        ' X,Y to screen coordinates!
        GetCursorPos tP
        ' Determine where to draw the splitter:
        If (m_eOrientation = espHorizontal) Then
            rcNew.Left = rcWindow.Left
            rcNew.Right = rcWindow.Right
            If (tP.Y >= rcWindow.Top) And (tP.Y < rcWindow.Bottom) Then
                rcNew.Top = tP.Y - 2
                rcNew.Bottom = tP.Y + 2
            Else
                If (tP.Y < rcWindow.Top) Then
                    rcNew.Top = rcWindow.Top - 2
                    rcNew.Bottom = rcNew.Top + 2
                Else
                    rcNew.Top = rcWindow.Bottom - 2
                    rcNew.Bottom = rcNew.Top + 2
                End If
            End If
        Else
            rcNew.Top = rcWindow.Top
            rcNew.Bottom = rcWindow.Bottom
            If (tP.x >= rcWindow.Left) And (tP.x <= rcWindow.Right) Then
                rcNew.Left = tP.x - 2
                rcNew.Right = tP.x + 2
            Else
                If (tP.x < rcWindow.Left) Then
                    rcNew.Left = rcWindow.Left - 2
                    rcNew.Right = rcNew.Left + 2
                Else
                    rcNew.Left = rcWindow.Right - 2
                    rcNew.Right = rcNew.Left + 2
                End If
            End If
        End If
        ' Draw the new rectangle
        Rectangle hDC, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom
        ' Store this position so we can erase it next time:
        LSet rcCurrent = rcNew
        ' Free the reference to the Desktop DC we got (make sure you do this!)
        DeleteDC hDC
        'DeleteObject hPen
    End If
End Sub

' changed to private
Private Function SplitterFormMouseUp(ByVal x As Long, ByVal Y As Long) As Boolean
    Dim hDC As Long
    Dim tP As POINTAPI
    Dim hWndClient As Long

    ' Release mouse capture:
    ReleaseCapture
    ' Don't leave orphaned rectangle on desktop; erase last rectangle.
    If (bDraw) Then
        bDraw = False
        ' Get the Desktop DC:
        hDC = CreateDCAsNull("DISPLAY", 0, 0, 0)
        ' Set to XOR drawing mode:
        SetROP2 hDC, R2_NOTXORPEN
        ' Erase the last rectangle:
        Rectangle hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
        ' Clear up the desktop DC:
        DeleteDC hDC
        ' Here we ensure the splitter is within bounds before releasing:
        GetCursorPos tP
        If (tP.x < rcWindow.Left) Then
            tP.x = rcWindow.Left
        End If
        If (tP.x > rcWindow.Right) Then
            tP.x = rcWindow.Right
        End If
        If (tP.Y < rcWindow.Top) Then
            tP.Y = rcWindow.Top
        End If
        If (tP.Y > rcWindow.Bottom) Then
            tP.Y = rcWindow.Bottom
        End If
        ' Move the splitter to the validated final position:
        If (m_eOrientation = espHorizontal) Then
            m_Offset = (tP.Y - m_Offset) * Screen.TwipsPerPixelY
        Else
            m_Offset = (tP.x - m_Offset) * Screen.TwipsPerPixelX
        End If
        ' Return true to tell the owner we have completed splitting:
        SplitterFormMouseUp = True
    End If
End Function

Private Sub Class_Initialize()
   m_eOrientation = espVertical
End Sub

Private Function ClassName(ByVal lhWnd As Long) As String
    Dim lLen As Long
    Dim sBuf As String
    
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lhWnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
End Function
'-- end code
