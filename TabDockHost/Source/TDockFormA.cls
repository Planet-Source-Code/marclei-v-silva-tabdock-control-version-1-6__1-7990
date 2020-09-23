VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************************************
' Project     : TabDock
' Module      : TDockForm
' Description : Save the form properties and information to be used
'               by the docking engine
' Created by  : Marclei V Silva
' Machine     : ZEUS
' Date-Time   : 11/05/2000 22:15:58
' Revisions   : See revisions.rtf shipped with this version
' ********************************************************************
Option Explicit

Private m_wndprcNext As Long
#If DEBUGWINDOWPROC Then
Private m_SCHook As WindowProcHook
#End If

' Keep up with the errors
Const g_ErrConstant As Long = vbObjectError + 1000
Const m_constClassName = "TDockForm"

Private m_lngErrNum As Long
Private m_strErrStr As String
Private m_strErrSource As String

Private m_bActive As Boolean
Private m_bStyleChanged As Boolean
Private m_bUnloaded As Boolean

' class properties
Private m_bDragging As Boolean      ' We are dragging?
Private m_SelectedItem As Integer   ' Host that was active during moving
Private m_lFloatingWidth As Long    ' Floating dimensions
Private m_lFloatingHeight As Long
Private m_lFloatingLeft As Long
Private m_lFloatingTop As Long
Private m_State As tdDockedState    ' form state dock/undocked
Private m_Style As tdDockStyles     ' availlable style
Private m_PanelPtr As Long          ' pointer to panel object
Private m_Key As String             ' key associatate to the form
Private m_hWnd As Long              ' handle for the form
Private m_FormPtr As Long           ' pointer to the form object
Private m_bRedraw As Boolean        ' flag that stop form painting or style change

Implements ISubclass

Public Property Get Object() As TDockForm
    Set Object = Me
End Property

Friend Property Get Extender() As Form
    Set Extender = ObjectFromPtr(m_FormPtr)
End Property

Friend Property Set Extender(ByVal vData As Form)
    m_FormPtr = PtrFromObject(vData)
    Attach vData.hWnd
End Property

Friend Function Move(Left, Optional Top, Optional Width, Optional Height)
    On Error Resume Next
    Extender.Move Left, Top, Width, Height
End Function

Friend Property Get Left() As Variant
    Left = Extender.Left
End Property

Friend Property Let Left(ByVal vData As Variant)
    Extender.Left = vData
End Property

Friend Property Get Top() As Variant
    Top = Extender.Top
End Property

Friend Property Let Top(ByVal vData As Variant)
    Extender.Top = vData
End Property

Friend Property Get Width() As Variant
    Width = Extender.Width
End Property

Friend Property Let Width(ByVal vData As Variant)
    Extender.Width = vData
End Property

Friend Property Get Height() As Variant
    Height = Extender.Height
End Property

Friend Property Let Height(ByVal vData As Variant)
    Extender.Height = vData
End Property

Friend Property Let Visible(ByVal vData As Boolean)
    Extender.Visible = vData
    If vData Then
        ' trigger TabDock.FormShow() Event
        Panel.Parent.TriggerEvent "ShowForm", Object
    Else
        ' trigger TabDock.Formhide() Event
        Panel.Parent.TriggerEvent "HideForm", Object
    End If
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns wether the form is visible or not"
    Visible = Extender.Visible
End Property

Public Property Let Style(ByVal vData As tdDockStyles)
Attribute Style.VB_Description = "Returns or sets the form style. The style may allow or avoid some features of a docked window. Check all the possible styles to choose the apropriate one."
    m_Style = vData
End Property

Public Property Get Style() As tdDockStyles
Attribute Style.VB_Description = "Returns or sets the form style. The style may allow or avoid some features of a docked window. Check all the possible styles to choose the apropriate one."
    Style = m_Style
End Property

Public Property Let Key(ByVal vData As String)
Attribute Key.VB_Description = "Returns the key of the form to be used with TabDock functions like FormShow() or FormHide. Instead of the key you may use the Form index to retrieve form reference"
    m_Key = vData
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key of the form to be used with TabDock functions like FormShow() or FormHide. Instead of the key you may use the Form index to retrieve form reference"
    Key = m_Key
End Property

Public Property Let FloatingHeight(ByVal vData As Long)
Attribute FloatingHeight.VB_Description = "Returns the floating height of the form"
    m_lFloatingHeight = vData
End Property

Public Property Get FloatingHeight() As Long
Attribute FloatingHeight.VB_Description = "Returns the floating height of the form"
    FloatingHeight = m_lFloatingHeight
End Property

Public Property Let FloatingWidth(ByVal vData As Long)
Attribute FloatingWidth.VB_Description = "Returns the floating Width of the form"
    m_lFloatingWidth = vData
End Property

Public Property Get FloatingWidth() As Long
Attribute FloatingWidth.VB_Description = "Returns the floating Width of the form"
    FloatingWidth = m_lFloatingWidth
End Property

Public Property Let FloatingLeft(ByVal vData As Long)
Attribute FloatingLeft.VB_Description = "Returns the floating Left of the form"
    m_lFloatingLeft = vData
End Property

Public Property Get FloatingLeft() As Long
Attribute FloatingLeft.VB_Description = "Returns the floating Left of the form"
    FloatingLeft = m_lFloatingLeft
End Property

Public Property Let FloatingTop(ByVal vData As Long)
Attribute FloatingTop.VB_Description = "Returns the floating Top of the form"
    m_lFloatingTop = vData
End Property

Public Property Get FloatingTop() As Long
Attribute FloatingTop.VB_Description = "Returns the floating Top of the form"
    FloatingTop = m_lFloatingTop
End Property

Friend Property Let Redraw(ByVal vData As Boolean)
    m_bRedraw = vData
End Property

Friend Property Get Redraw() As Boolean
    Redraw = m_bRedraw
End Property

Friend Property Set Panel(ByVal vData As TTabDockHost)
    m_PanelPtr = PtrFromObject(vData)
End Property

Public Property Get Panel() As TTabDockHost
Attribute Panel.VB_Description = "Returns the panel that is responsible for the form. This property changes when you dock the form into a different panel."
    Set Panel = ObjectFromPtr(m_PanelPtr)
End Property

Public Property Let State(ByVal vData As tdDockedState)
Attribute State.VB_Description = "Returns the state of the form. Docked or Undocked."
    m_State = vData
End Property

Public Property Get State() As tdDockedState
Attribute State.VB_Description = "Returns the state of the form. Docked or Undocked."
    State = m_State
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the form handle. This may not be used in api callings once the docked form is already subclassed and another subclassing of this form may crash Vb IDE or unpredictable results may take place."
    hWnd = Extender.hWnd
End Property

' ******************************************************************************
' Routine       : (Sub) Class_Initialize
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:36:23
' Inputs        : N/A
' Outputs       : N/A
' Modifications : N/A
' Description   : Initialize class
' ******************************************************************************
Private Sub Class_Initialize()
    On Error GoTo Err_Class_Initialize
    Const constSource As String = m_constClassName & ".Class_Initialize"

    m_bActive = False
    m_bStyleChanged = False
    m_SelectedItem = 0
    
    Exit Sub
Err_Class_Initialize:
    Err.Raise Description:=Err.Description, _
       Number:=Err.Number, _
       Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Class_Terminate
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:36:03
' Inputs        : N/A
' Outputs       : N/A
' Modifications : N/A
' Description   : Terminate the class (clean it up)
' ******************************************************************************
Private Sub Class_Terminate()
    On Error GoTo Err_Class_Terminate
    Const constSource As String = m_constClassName & ".Class_Terminate"
    
    ' detach all window messages
    Detach
    ' try to unload the extender
    Unload Extender
    
    Exit Sub
Err_Class_Terminate:
    Err.Raise Description:=Err.Description, _
       Number:=Err.Number, _
       Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) DockForm_StoreFloatingPos
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 09/06/2000 12:47:34
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Store the floating position of the form
' ******************************************************************************
Private Sub DockForm_StoreFloatingPos(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
    m_lFloatingLeft = lLeft
    m_lFloatingTop = lTop
    m_lFloatingWidth = lWidth
    m_lFloatingHeight = lHeight
End Sub

' ******************************************************************************
' Routine       : (Sub) DockForm_Show
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:35:36
' Inputs        : N/A
' Outputs       : N/A
' Modifications : N/A
' Description   : shows the docked form no matter it is docked or undocked
' ******************************************************************************
Friend Sub DockForm_Show()
    On Error GoTo Err_DockForm_Show
    Const constSource As String = m_constClassName & ".DockForm_Show"

    If State = tdDocked Then
        Panel.Dock Me
    Else
        Panel.UnDock Me
    End If
    m_bUnloaded = False
    DockForm_ResetStyle
    DockForm_Paint
    
    Exit Sub
Err_DockForm_Show:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Hide
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:35:19
' Inputs        : N/A
' Outputs       : N/A
' Modifications : N/A
' Description   : hides the docked form no matter it is docked or undocked
' ******************************************************************************
Friend Sub DockForm_Hide()
    On Error GoTo Err_Hide
    Const constSource As String = m_constClassName & ".Hide"

    Visible = False
    If Panel.IsVisible Then
        Panel.Repaint
    Else
        Panel.Visible = False
    End If

    Exit Sub
Err_Hide:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : Dock
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:22:12
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : dock the form
' ******************************************************************************
Public Sub DockForm_Dock()
    Panel.Dock Me
End Sub

' ******************************************************************************
' Routine       : UnDock
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:22:26
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : undock the form
' ******************************************************************************
Public Sub DockForm_UnDock()
    Panel.UnDock Me
End Sub

Private Sub DockForm_Activate(ByVal bByMouse As Boolean)
    DockForm_Paint
End Sub

Private Sub DockForm_Paint()
    Dim Rc As RECT
    Dim bdrStyle As Long
    Dim bdrSides As Long
    Dim BorderStyle  As tdBorderStyles
    Dim hDC As Long
    
    If Me.State = tdUndocked Or _
        m_bRedraw = False Or _
        Extender.Visible = False Then
        Exit Sub
    End If
    ' draw a custom border based on parante's color
    DrawBorder Extender, Panel.Parent.BackColor
    ' retrieve TabDock border style
    BorderStyle = Panel.Parent.BorderStyle
    ' all sides must be updated
    bdrSides = BF_RECT
    ' update border styles
    If BorderStyle = bdrFlat Then bdrSides = bdrSides Or BF_FLAT
    If BorderStyle = bdrMono Then bdrSides = bdrSides Or BF_MONO
    If BorderStyle = bdrSoft Then bdrSides = bdrSides Or BF_SOFT
    Select Case BorderStyle
        Case bdrRaisedOuter: bdrStyle = BDR_RAISEDOUTER
        Case bdrRaisedInner: bdrStyle = BDR_RAISEDINNER
        Case bdrRaised: bdrStyle = EDGE_RAISED
        Case bdrSunkenOuter: bdrStyle = BDR_SUNKENOUTER
        Case bdrSunkenInner: bdrStyle = BDR_SUNKENINNER
        Case bdrSunken: bdrStyle = EDGE_SUNKEN
        Case bdrEtched: bdrStyle = EDGE_ETCHED
        Case bdrBump: bdrStyle = EDGE_BUMP
        Case bdrFlat: bdrStyle = BDR_SUNKEN
        Case bdrMono: bdrStyle = BDR_SUNKEN
        Case bdrSoft: bdrStyle = BDR_RAISED
    End Select
    ' get a window rect by hand
    ' GetWindowRect will not work here!
    Rc.Left = 0
    Rc.Top = 0
    Rc.Bottom = Extender.Height / Screen.TwipsPerPixelY
    Rc.Right = Extender.Width / Screen.TwipsPerPixelY
    ' First get the window DC
    hDC = GetWindowDC(hWnd)
    ' Simply call the API and draw the edge.
    DrawEdge hDC, Rc, bdrStyle, bdrSides
    ' release it
    ReleaseDC hWnd, hDC
End Sub

Private Sub DockForm_DblClick()
    If State = tdDocked Then
        Panel.UnDock Me
    Else
        Panel.Dock Me
    End If
End Sub

' ******************************************************************************
' Routine       : DockForm_MouseDown
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:23:00
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Do some actions before dragging a form
' ******************************************************************************
Private Sub DockForm_MouseDown(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long, Cancel As Boolean)
    If Style And tdDockFloat Then
        m_bDragging = True
        ' if form is undocked then save the floating position
        If State = tdUndocked Then
            DockForm_StoreFloatingPos FormLeft, FormTop, FormWidth, FormHeight
        End If
    Else
        Cancel = True
    End If
End Sub

' ******************************************************************************
' Routine       : DockForm_MouseUp
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:23:29
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Actions after a form is dropped somewhere
' ******************************************************************************
Private Sub DockForm_MouseUp(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
    ' If no host was found to dock the form then
    ' undock it if it is not undocked
    If m_SelectedItem = 0 Then
        If m_State = tdDocked Then
            ' undock the form only if it can float
            If m_Style And tdDockFloat Then
                Panel.UnDock Me
                ' repaint the host
                Panel.Repaint
            End If
        End If
        Exit Sub
    End If
    ' if form is undocked then dock it
    If m_State = tdUndocked Then
        ' dock the form
        If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
            Panel.Parent.Panels(m_SelectedItem).Dock Me
        End If
    ElseIf m_State = tdDocked Then
        ' if the selected host is the active host
        ' just repaint the host panel
        If m_SelectedItem = Panel.Index Then
            Panel.DockSwap Me, FormLeft, FormTop
            Panel.Repaint
        Else
            ' if selected host is different from
            ' active host, first undock the form
            If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
                Panel.UnDock Me
                ' dock the form into the selected host
                Panel.Parent.Panels(m_SelectedItem).Dock Me
            End If
        End If
    End If
    ' A docking site changed then we must set
    ' the new dock host for this form
    If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
        Set Panel = Panel.Parent.Panels(m_SelectedItem)
    End If
End Sub

' ******************************************************************************
' Routine       : DockForm_MouseMove
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:25:03
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Called when the form is moving around
' ******************************************************************************
Private Sub DockForm_MouseMove(ByRef FormLeft As Long, ByRef FormTop As Long, ByRef FormWidth As Long, ByRef FormHeight As Long, ByRef lBorderWidth As Long)
    Dim Rc As RECT
    Dim I As Integer
    Dim Found As Boolean
    Dim pt As POINTAPI
    Dim lX1Offset As Long
    Dim lX2Offset As Long
    Dim lYOffset As Long
    Dim lObjWidth As Long
    Dim lObjHeight As Long
    
    ' flag that indicates we found a dock host
    Found = False
    ' get the cursor position
    GetCursorPos pt
    ' loop the panels to see we find a docking host
    ' for this form
    For I = 1 To 4
        ' get the host rect
        Panel.Parent.Panels(I).GetHostRect Rc
        ' check if the form is within the panel's boundaries
        If PtInRect(Rc, pt.x, pt.Y) Then
            If Panel.Parent.Panels(I).AllowDocking(Me) Then
                ' a host was found and it is different from the
                ' the current selected one
                If Not Panel.Parent.Panels(I).Index = m_SelectedItem Then
                    ' host selection changed here
                    m_SelectedItem = Panel.Parent.Panels(I).Index
                    ' we don't have a dragging boundary set yet
                    ' for this host then update with the host's rect
                    lObjWidth = Rc.Right - Rc.Left
                    lObjHeight = Rc.Bottom - Rc.Top
                    If Panel.Parent.Panels(I).Align = tdAlignLeft Or Panel.Parent.Panels(I).Align = tdAlignRight Then
                        lX1Offset = pt.x - FormLeft
                        lX2Offset = lObjWidth * lX1Offset / FormWidth 'pt.Y - rc.Top
                        FormLeft = pt.x - lX2Offset
                        FormWidth = lObjWidth
                        FormHeight = lObjHeight
                    Else
                        lX1Offset = pt.Y - FormTop
                        lX2Offset = lObjHeight * lX1Offset / FormHeight 'pt.Y - rc.Top
                        FormTop = pt.Y - lX2Offset
                        FormWidth = lObjWidth
                        FormHeight = lObjHeight
                    End If
                End If
                ' change border width (special effect)
                lBorderWidth = 2
                ' we found a host, set flag = true
                Found = True
            End If
        End If
        ' we have found a host so exit
        If Found Then Exit For
    Next
    ' if the form has not found any host to dock in
    ' then let the rectangle image the same as its
    ' floating original rectangle
    If Not Found Then
        ' if we are dragging and the form is docked
        If m_bDragging And m_State = tdDocked Then
            ' get active host rect
            GetWindowRect Panel.hWnd, Rc
            ' restore form left property to its
            ' floating left position
            lObjWidth = Rc.Right - Rc.Left
            lObjHeight = Rc.Bottom - Rc.Top
            lX1Offset = pt.x - FormLeft
            lX2Offset = m_lFloatingWidth * lX1Offset / lObjWidth 'pt.Y - rc.Top
            FormLeft = pt.x - lX2Offset
            ' flag must be false
            m_bDragging = False
        End If
        ' change border width (just a special effect)
        lBorderWidth = 3
        ' update width and height with floating dimensions
        FormWidth = m_lFloatingWidth
        FormHeight = m_lFloatingHeight
        ' store floating position
        DockForm_StoreFloatingPos FormLeft, FormTop, FormWidth, FormHeight
        ' clear host selection
        m_SelectedItem = 0
    End If
End Sub

' ******************************************************************************
' Routine       : DockForm_Unload
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:27:09
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Hides the form
' ******************************************************************************
Private Sub DockForm_Unload()
    DockForm_Hide
End Sub

' ******************************************************************************
' Routine       : DockForm_ResetStyle
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:27:39
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Resets the style to a tollwindow
'                 This routine is necessay when the style of
'                 the form changes to force it to be a toolwindow
' ******************************************************************************
Private Sub DockForm_ResetStyle()
    Dim New_Style As Long
    
    ' if redraw is disable do not set any style
    ' or else you may get a loop here
    If m_bRedraw = False Then Exit Sub
    ' set this flag to true so that the next set style event will
    ' not be handled (trick!)
    m_bStyleChanged = True
    ' get current window style
    New_Style = GetWindowLong(hWnd, GWL_STYLE)
    ' check form state
    If State = tdDocked Then
        New_Style = New_Style And Not WS_THICKFRAME
        SetWindowLong hWnd, GWL_STYLE, New_Style
    Else
        New_Style = New_Style And Not WS_DLGFRAME
        New_Style = New_Style Or WS_CAPTION
        New_Style = New_Style Or WS_THICKFRAME
    End If
    ' get extended style
    New_Style = GetWindowLong(hWnd, GWL_EXSTYLE)
    ' set it to toolwindow
    New_Style = New_Style Or WS_EX_TOOLWINDOW
    New_Style = New_Style And Not (WS_EX_APPWINDOW)
    SetWindowLong hWnd, GWL_EXSTYLE, New_Style
    ' reset flag here
    m_bStyleChanged = False
End Sub

' ******************************************************************************
' Routine       : DockForm_Deactivate
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:29:52
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Called when the form is deactivated
' ******************************************************************************
Private Sub DockForm_Deactivate()
    DockForm_Paint
End Sub


' ******************************************************************************
' Routine       : (Sub) DragObject
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:38:15
' Inputs        :
' Outputs       :
' Modifications :
' Credits       : This rotine was extract from the DockForm_Dragger project
'               : by Steve
'               : Steve81@ mediaone.net
'               : http://www.vbtutor.com
'               : Thanks Steve!
'               :
' Modifications : Added the ability to change the dragging border width
'               : Removed reposition
'               :
' Description   : Procedure which simulates windows dragging of an object
' ******************************************************************************
Private Sub DragObject(ByVal hWnd As Long)
    Dim pt As POINTAPI
    Dim ptPrev As POINTAPI
    Dim objRect As RECT
    Dim DragRect As RECT
    Dim na As Long
    Dim lBorderWidth As Long
    Dim lObjTop As Long
    Dim lObjLeft As Long
    Dim lObjWidth As Long
    Dim lObjHeight As Long
    Dim lXOffset As Long
    Dim lYOffset As Long
    Dim bMoved As Boolean
    Dim bCancel As Boolean
    
    ReleaseCapture
    GetWindowRect hWnd, objRect
    lObjWidth = objRect.Right - objRect.Left
    lObjHeight = objRect.Bottom - objRect.Top
    GetCursorPos pt
    'Store the initial cursor position
    ptPrev.x = pt.x
    ptPrev.Y = pt.Y
    'Set the initial rectangle, and draw it to show the user that
    'the object can be moved
    lXOffset = pt.x - objRect.Left
    lYOffset = pt.Y - objRect.Top
    With DragRect
        .Left = pt.x - lXOffset
        .Top = pt.Y - lYOffset
        .Right = .Left + lObjWidth
        .Bottom = .Top + lObjHeight
    End With
    ' use form border width highlighting
    lBorderWidth = 3
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
    ' Move the object
    DockForm_MouseDown pt.x - lXOffset, pt.Y - lYOffset, lObjWidth, lObjHeight, bCancel
    If bCancel = False Then
        Do While GetKeyState(VK_LBUTTON) < 0
            'Debug.Print "moving..."
            GetCursorPos pt
            If pt.x <> ptPrev.x Or pt.Y <> ptPrev.Y Then
                ptPrev.x = pt.x
                ptPrev.Y = pt.Y
                ' erase the previous drag rectangle if any
                DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                'Tell the user we've moved
                lObjLeft = pt.x - lXOffset
                lObjTop = pt.Y - lYOffset
                DockForm_MouseMove lObjLeft, lObjTop, lObjWidth, lObjHeight, lBorderWidth
                ' Adjust the height/width
                With DragRect
                    .Left = lObjLeft
                    .Top = lObjTop
                    .Right = .Left + lObjWidth
                    .Bottom = .Top + lObjHeight
                End With
                DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                bMoved = True
                lXOffset = pt.x - DragRect.Left
                lYOffset = pt.Y - DragRect.Top
            End If
            DoEvents
        Loop
    End If
    ' erase the previous drag rectangle if any
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
    ' move and repaint the window
    If bMoved Then
        ' If m_RepositionForm Then
        MoveWindow hWnd, DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
        ' End If
        ' tell the user we've dropped the form
        DockForm_MouseUp DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top
    End If
End Sub

Private Sub DockForm_CaptionClick(ByVal Button As Integer, ByVal x As Single, ByVal Y As Single)
    ' trigger event
    Panel.Parent.TriggerEvent "CaptionClick", Object, Button, x * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    ' When I trigger this event the program locks up
    ' then I discovered that when we send the focus to the main window
    ' the problem is solved. Why? I really don't know
    ' I'm not so good at API the way you may think
    ' but for now it is working fine...
    Putfocus Panel.Parent.Parent.hWnd
End Sub

' ******************************************************************************
' Routine       : (Sub) DrawDragRectangle
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:40:49
' Inputs        :
' Outputs       :
' Modifications :
' Credits       : This rotine was extract from the DockForm_Dragger project
'                 by Steve (Steve81@ mediaone.net)
'                 http://www.vbtutor.com
'                 Thanks Steve!
' Modifications : Changed the pen color to vbButtonFace (that's all)
'
' Description   : Draw a rectangle using the Win32 API
' ******************************************************************************
Private Sub DrawDragRectangle(ByVal x As Long, ByVal Y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long)
    Dim hDC As Long
    Dim hPen As Long
    
    hPen = CreatePen(PS_INSIDEFRAME, lWidth, vbButtonFace) '&H0) '&HE0E0E0)
    hDC = GetDC(0)
    Call SelectObject(hDC, hPen)
    Call SetROP2(hDC, R2_NOTXORPEN)
    Call Rectangle(hDC, x, Y, X1, Y1)
    Call SelectObject(hDC, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hDC, hPen)
    Call ReleaseDC(0, hDC)
End Sub

' ******************************************************************************
' Routine       : Attach
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:19:32
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : attach window messages to a specific hWnd
' ******************************************************************************
Friend Sub Attach(ByVal hWndA As Long)
    Detach
    m_hWnd = hWndA
    AttachMessage Me, m_hWnd, WM_DESTROY
    AttachMessage Me, m_hWnd, WM_ACTIVATE
    AttachMessage Me, m_hWnd, WM_SIZE
    AttachMessage Me, m_hWnd, WM_MOVE
    AttachMessage Me, m_hWnd, WM_SYSCOMMAND
    AttachMessage Me, m_hWnd, WM_STYLECHANGED
    AttachMessage Me, m_hWnd, WM_CLOSE
    AttachMessage Me, m_hWnd, WM_CONTEXTMENU
    AttachMessage Me, m_hWnd, WM_NCRBUTTONDOWN
    AttachMessage Me, m_hWnd, WM_PAINT
'    AttachMessage Me, m_hWnd, WM_NCLBUTTONDOWN
    AttachMessage Me, m_hWnd, WM_NCLBUTTONDBLCLK
End Sub

' ******************************************************************************
' Routine       : Detach
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:19:55
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Detach window messages
' ******************************************************************************
Friend Sub Detach()
    If Not m_hWnd = 0 Then
        DetachMessage Me, m_hWnd, WM_DESTROY
        DetachMessage Me, m_hWnd, WM_SIZE
        DetachMessage Me, m_hWnd, WM_ACTIVATE
        DetachMessage Me, m_hWnd, WM_MOVE
        DetachMessage Me, m_hWnd, WM_SYSCOMMAND
        DetachMessage Me, m_hWnd, WM_STYLECHANGED
        DetachMessage Me, m_hWnd, WM_CLOSE
        DetachMessage Me, m_hWnd, WM_CONTEXTMENU
        DetachMessage Me, m_hWnd, WM_NCRBUTTONDOWN
        DetachMessage Me, m_hWnd, WM_PAINT
'        DetachMessage Me, m_hWnd, WM_NCLBUTTONDOWN
        DetachMessage Me, m_hWnd, WM_NCLBUTTONDBLCLK
    End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
    '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
    ' Process before windows:
    Select Case CurrentMessage
        Case WM_SYSCOMMAND
            ISubclass_MsgResponse = emrConsume
        Case WM_CONTEXTMENU
            ISubclass_MsgResponse = emrConsume
        Case WM_NCRBUTTONDOWN
            ISubclass_MsgResponse = emrConsume
'        Case WM_NCLBUTTONDOWN
'            ISubclass_MsgResponse = emrPreprocess
'        Case WM_ENTERSIZEMOVE
'            ISubclass_MsgResponse = emrPostProcess
'        Case WM_EXITSIZEMOVE
'            ISubclass_MsgResponse = emrConsume
        Case WM_SIZE
            ISubclass_MsgResponse = emrPreprocess
        Case WM_MOVE
            ISubclass_MsgResponse = emrConsume
        Case WM_PAINT
            ISubclass_MsgResponse = emrPreprocess
        Case Else
            ISubclass_MsgResponse = emrPostProcess
    End Select
End Property

' ******************************************************************************
' Routine       : ISubClass_WindowProc
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:30:36
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Window messages handler for the form
' ******************************************************************************
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pt As POINTAPI
    Dim uCmdType As Long
    
    Select Case iMsg
        Case WM_SYSCOMMAND
            Debug.Print "SYSCOMMAND"
            uCmdType = wParam And &HFFF0
            If uCmdType = SC_MOVE Then
                ReleaseCapture
                DragObject hWnd
            ElseIf uCmdType = SC_CLOSE Then
                m_bUnloaded = True
                DockForm_Unload
            Else
                ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            End If
        Case WM_NCLBUTTONDBLCLK
            Debug.Print "NCLBUTTONDBLCLK"
            If wParam = HTCAPTION Then
                DockForm_DblClick
            End If
'        Case WM_NCLBUTTONDOWN
'            Debug.Print "NCLBUTTONDOWN"
'            If (wParam = HTCAPTION) Then
'                Debug.Print "NCLBUTTONDOWN::CAPTION"
'            ElseIf wParam = HTCLOSE Then
'                m_bUnloaded = True
'                DockForm_Unload
'            End If
        Case WM_STYLECHANGED
            If m_bStyleChanged = False Then
                DockForm_ResetStyle
            End If
        Case WM_ACTIVATE
            ' Form is activated/deactivated:
            If wParam = 0 Then
                Debug.Print "DEACT"
                m_bActive = False
                DockForm_Deactivate
            Else
                Debug.Print "ACT"
                m_bActive = True
                DockForm_Activate wParam = 2
            End If
        Case WM_NCRBUTTONDOWN
            Debug.Print "NCRBUTTONDOWN"
            If (wParam = HTCAPTION) Then
                DockForm_CaptionClick vbRightButton, LoWord(lParam), HiWord(lParam)
            End If
        Case WM_CONTEXTMENU
            Debug.Print "WM_CONTEXTMENU"
        Case WM_CLOSE
            Debug.Print "WM_CLOSE"
        Case WM_DESTROY
            Debug.Print "WM_DESTROY"
            Detach
        Case WM_PAINT
            Debug.Print "WM_PAINT"
            DockForm_Paint
        Case WM_MOVE
            Debug.Print "WM_MOVE"
            DockForm_Paint
        Case WM_SIZE
            Debug.Print "WM_SIZE"
            DockForm_Paint
    End Select
End Function
'-- end code
