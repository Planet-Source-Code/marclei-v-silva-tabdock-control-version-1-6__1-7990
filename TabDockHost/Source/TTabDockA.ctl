VERSION 5.00
Begin VB.UserControl TTabDock 
   CanGetFocus     =   0   'False
   ClientHeight    =   528
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   588
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "TTabDockA.ctx":0000
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   49
   ToolboxBitmap   =   "TTabDockA.ctx":0842
End
Attribute VB_Name = "TTabDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' ******************************************************************************
' Control    : TabDock.ctl
' Created by : Marclei V Silva
' Machine    : ZEUS
' Date-Time  : 09/05/2000 3:13:22
' Description: Docking system engine
' ******************************************************************************
Option Explicit
Option Base 1

' Keep up with the errors
Const g_ErrConstant As Long = vbObjectError + 1000
Const m_constClassName = "TTabDock"

Private m_lngErrNum As Long
Private m_strErrStr As String
Private m_strErrSource As String

Private m_Panels As TTabDockHosts
Private m_DockedForms As TDockForms

' Events Held by this control
Public Event FormDocked(ByVal DockedForm As TDockForm)
Attribute FormDocked.VB_Description = "Occurs when the user drag and dock a form at a specific panel on the screen"
Public Event FormUnDocked(ByVal DockedForm As TDockForm)
Attribute FormUnDocked.VB_Description = "Occurs when the user undocks a form from a specific panel"
Public Event FormShow(ByVal DockedForm As TDockForm)
Attribute FormShow.VB_Description = "Occurs when a form is shown in the screen. This event accurs no matter the form is docked or undocked"
Public Event FormHide(ByVal DockedForm As TDockForm)
Attribute FormHide.VB_Description = "Hides a specific form"
Public Event MenuClick(ByVal ItemIndex As Long)
Public Event PanelResize(ByVal Panel As TTabDockHost)
Attribute PanelResize.VB_Description = "Occurs when a specific panel is resized. This is useful when you want to set a specific Height or width for a panel in the screen or avoid user to resize a panel to a not desired size."
Public Event PanelClick(ByVal Panel As TTabDockHost)
Public Event CaptionClick(ByVal DockedForm As TDockForm, ByVal Button As Integer, ByVal x As Single, ByVal Y As Single)
Attribute CaptionClick.VB_Description = "Occurs when the user clicks on the caption bar of a form. This is very useful when we want to show a popup menu for that form like Dockable or Hide."

' Default Property Values:
Const m_def_BackColor = &H8000000F
Const m_def_BorderStyle = 0 ' flat
Const m_def_PanelHeight = 1300
Const m_def_PanelWidth = 2500
Const m_def_Visible = 0

' Property Variables:
Private m_BackColor As OLE_COLOR
Private m_BorderStyle As tdBorderStyles
Private m_Parent As Object
Private m_PanelHeight As Long
Private m_PanelWidth As Long
Private m_Visible As Boolean
Private m_bLoaded As Boolean

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,0
Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Show/Hides the docking system frame"
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal New_Visible As Boolean)
Attribute Visible.VB_Description = "Show/Hides the docking system frame"
    If Ambient.UserMode = False Then Err.Raise 387
    m_Visible = New_Visible
    PropertyChanged "Visible"
    LockWindowUpdate Extender.Parent.hWnd
        m_Panels(tdAlignLeft).Visible = New_Visible
        m_Panels(tdAlignRight).Visible = New_Visible
        m_Panels(tdAlignTop).Visible = New_Visible
        m_Panels(tdAlignBottom).Visible = New_Visible
    LockWindowUpdate ByVal 0&
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Generally this is the MDI form the control was dropped in"
    Set Parent = Extender.Parent
End Property

Property Get Panels() As TTabDockHosts
Attribute Panels.VB_Description = "Panels of the docking system"
    Set Panels = m_Panels
End Property

Property Get DockedForms() As TDockForms
Attribute DockedForms.VB_Description = "Collection of docked forms"
    Set DockedForms = m_DockedForms
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the back color of the docking frame"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Returns or sets the back color of the docking frame"
    Dim I As Integer
    
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    LockWindowUpdate Extender.Parent.hWnd
        For I = 1 To Panels.Count
            Panels(I).BackColor = New_BackColor
        Next
    LockWindowUpdate ByVal 0&
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get BorderStyle() As tdBorderStyles
Attribute BorderStyle.VB_Description = "Returns or set the border style of the docked forms."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As tdBorderStyles)
Attribute BorderStyle.VB_Description = "Returns or set the border style of the docked forms."
    Dim I As Integer
    
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    LockWindowUpdate Extender.Parent.hWnd
    For I = 1 To Panels.Count
        Panels(I).DockArrange
    Next
    LockWindowUpdate ByVal 0&
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,0,2100
Public Property Get PanelHeight() As Long
Attribute PanelHeight.VB_Description = "Returns or sets the initial height of top and bottom panels"
    PanelHeight = m_PanelHeight
End Property

Public Property Let PanelHeight(ByVal New_PanelHeight As Long)
Attribute PanelHeight.VB_Description = "Returns or sets the initial height of top and bottom panels"
    If Ambient.UserMode Then Err.Raise 382
    m_PanelHeight = New_PanelHeight
    PropertyChanged "PanelHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,0,1000
Public Property Get PanelWidth() As Long
Attribute PanelWidth.VB_Description = "Returns or sets a initial Width for the left and right panels"
    PanelWidth = m_PanelWidth
End Property

Public Property Let PanelWidth(ByVal New_PanelWidth As Long)
Attribute PanelWidth.VB_Description = "Returns or sets a initial Width for the left and right panels"
    If Ambient.UserMode Then Err.Raise 382
    m_PanelWidth = New_PanelWidth
    PropertyChanged "PanelWidth"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_BorderStyle = m_def_BorderStyle
    m_PanelHeight = m_def_PanelHeight
    m_PanelWidth = m_def_PanelWidth
    m_Visible = m_def_Visible
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_PanelHeight = PropBag.ReadProperty("PanelHeight", m_def_PanelHeight)
    m_PanelWidth = PropBag.ReadProperty("PanelWidth", m_def_PanelWidth)
    m_Visible = PropBag.ReadProperty("Visible", m_def_Visible)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("PanelHeight", m_PanelHeight, m_def_PanelHeight)
    Call PropBag.WriteProperty("PanelWidth", m_PanelWidth, m_def_PanelWidth)
    Call PropBag.WriteProperty("Visible", m_Visible, m_def_Visible)
End Sub

Private Sub UserControl_Initialize()
On Error GoTo Err_UserControl_Initialize
    Const constSource As String = m_constClassName & ".UserControl_Initialize"

    Set m_DockedForms = New TDockForms
    Set m_Panels = New TTabDockHosts
    
Exit Sub
Err_UserControl_Initialize:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Err_UserControl_Terminate
    Const constSource As String = m_constClassName & ".UserControl_Terminate"
    
    Set m_Panels = Nothing
    Set m_DockedForms = Nothing
    
Exit Sub
Err_UserControl_Terminate:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub UserControl_Paint()
On Error GoTo Err_UserControl_Paint
    Const constSource As String = m_constClassName & ".UserControl_Paint"

    Dim Edge As RECT                                ' Rectangle edge of control
    
    Edge.Left = 0                                   ' Set rect edges to outer
    Edge.Top = 0                                    ' most position in pixels
    Edge.Bottom = ScaleHeight
    Edge.Right = ScaleWidth
    DrawEdge hDC, Edge, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT ' Draw Edge...

Exit Sub
Err_UserControl_Paint:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub UserControl_Resize()
On Error GoTo Err_UserControl_Resize
    Const constSource As String = m_constClassName & ".UserControl_Resize"

    ' set the control to 32 pixels wide
    UserControl.Width = 32 * Screen.TwipsPerPixelX
    UserControl.Height = 32 * Screen.TwipsPerPixelY

Exit Sub
Err_UserControl_Resize:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : AddForm
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/006:00:45
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Adds forms to the main engine
' ******************************************************************************
Public Function AddForm(ByVal Item As Object, Optional State As tdDockedState = tdUndocked, Optional Align As tdAlignProperty = tdAlignLeft, Optional Key As String, Optional Style As tdDockStyles) As TDockForm
Attribute AddForm.VB_Description = "Add a form reference to the dock system and updates its initial properties"
On Error GoTo Err_AddForm
    Const constSource As String = m_constClassName & ".AddForm"
    
    If IsFormLoaded(Item.hWnd) Then
        m_strErrStr = "Form is already loaded"
        m_strErrSource = constSource
        m_lngErrNum = 0
        m_lngErrNum = m_lngErrNum + g_ErrConstant
        Err.Raise Description:="Unexpected Error: " & m_strErrStr, _
                  Number:=m_lngErrNum, _
                  Source:=constSource
    End If
    ' if we are initializing (panels were not created) then create panels
    If m_bLoaded = False Then
        LoadPanels
    End If
    ' loads the form if it wasn't loaded yet!
    Load Item
    ' if the form style was not furnished then set
    ' all styles available to the form
    If IsMissing(Style) Or IsEmpty(Style) Or Style = 0 Or Style = tdShowInvisible Then
        Style = Style Or tdDockFloat
        Style = Style Or tdDockLeft
        Style = Style Or tdDockRight
        Style = Style Or tdDockTop
        Style = Style Or tdDockBottom
    End If
    ' add the form to the list
    Set AddForm = m_DockedForms.Add(Item, Panels(Align), Style, State, Key)

Exit Function
Err_AddForm:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Function

' ******************************************************************************
' Routine       : (Sub) LoadPanels
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 12/06/2000 - 22:11:07
' Inputs        : N/A
' Outputs       : N/A
' Modifications :
' Description   : Load the panels for the docking system
' ******************************************************************************
Private Sub LoadPanels()
On Error GoTo Err_LoadPanels
    Const constSource As String = m_constClassName & ".LoadPanels"

    Dim I As Integer
    Dim pict As VB.PictureBox
    
    ' only to avoid panels re-loading
    If m_bLoaded = True Then Exit Sub
    ' loop to create the 4 panels (left, top, right, bottom panels)
    For I = 1 To 4
        ' add a picture box at run-time to the extender (form)
        Set pict = LoadControl(Extender.Parent, "VB.PictureBox", "Host" & CStr(I))
        pict.BackColor = m_BackColor
        ' add a new panel to the list, the container
        ' will be our picture box
        m_Panels.Add I, m_PanelHeight, m_PanelWidth, False, Me, pict, "Host" & CStr(I)
    Next
    m_bLoaded = True
    
Exit Sub
Err_LoadPanels:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Function) LoadControl
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 12/06/2000 - 22:22:42
' Inputs        :
' Outputs       :
' Credits       : This code was extract from
'                 FreeVBCode.com (http://www.freevbcode.com)
' Modifications :
' Description   : Load a form control at run-time
' ******************************************************************************
Private Function LoadControl(oForm As Object, CtlType As String, ctlName As String, Optional CtlContainer) As Object
    Dim oCtl As Object
    On Error Resume Next
    If IsObject(oForm.Controls) Then
        If IsMissing(CtlContainer) Then
            Set oCtl = oForm.Controls.Add(CtlType, ctlName)
        Else
            Set oCtl = oForm.Controls.Add(CtlType, ctlName, CtlContainer)
        End If
        If Not oCtl Is Nothing Then
            Set LoadControl = oCtl
            Set oCtl = Nothing
        End If
    End If
End Function

' ******************************************************************************
' Routine       : (Sub) Show
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 12/06/2000 - 22:22:13
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Show panels and forms docked/undocked
' ******************************************************************************
Public Sub Show()
Attribute Show.VB_Description = "Show the host panels and update docked forms"
On Error GoTo Err_Show
    Const constSource As String = m_constClassName & ".Show"

    Dim I As Integer
    
    ' let's avoid some flickering...
    LockWindowUpdate Extender.Parent.hWnd
        ' dock/undock the forms
        For I = 1 To m_DockedForms.Count
            If (m_DockedForms(I).Style And tdShowInvisible) = False Then
                ' it it it is docked then dock it
                If m_DockedForms(I).State = tdDocked Then
                    m_DockedForms(I).Panel.Dock m_DockedForms(I)
                Else
                    ' just show
                    m_DockedForms(I).Panel.UnDock m_DockedForms(I)
                End If
            End If
        Next
    ' free willy! (I mean windows!)
    LockWindowUpdate 0

Exit Sub
Err_Show:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) TriggerEvent
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 12/06/2000 - 22:20:56
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Used to raise events to the form user
' ******************************************************************************
Friend Sub TriggerEvent(ByVal RaisedEvent As String, ParamArray aParams())
On Error GoTo Err_TriggerEvent
    Const constSource As String = m_constClassName & ".TriggerEvent"
    
    Select Case RaisedEvent
    Case "Dock"
        RaiseEvent FormDocked(aParams(0))
    Case "UnDock"
        RaiseEvent FormUnDocked(aParams(0))
    Case "ShowForm"
        RaiseEvent FormShow(aParams(0))
    Case "HideForm"
        RaiseEvent FormHide(aParams(0))
    Case "ResizePanel"
        RaiseEvent PanelResize(aParams(0))
    Case "MenuClick"
        RaiseEvent MenuClick(aParams(0))
    Case "PanelClick"
        RaiseEvent PanelClick(aParams(0))
    Case "CaptionClick"
        RaiseEvent CaptionClick(aParams(0), aParams(1), aParams(2), aParams(3))
    Case Else
        Debug.Print "Event no handled: " & RaisedEvent
    End Select
    
Exit Sub
Err_TriggerEvent:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Function IsFormLoaded(hWndA As Long) As Boolean
On Error GoTo Err_IsFormLoaded
    Const constSource As String = m_constClassName & ".IsFormLoaded"

    Dim I As Integer
    For I = 1 To m_DockedForms.Count
        If m_DockedForms(I).hWnd = hWndA Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
    IsFormLoaded = False

Exit Function
Err_IsFormLoaded:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Function

Public Sub FormShow(Index As Variant)
Attribute FormShow.VB_Description = "Shows a docked form"
On Error GoTo Err_FormShow
    Const constSource As String = m_constClassName & ".FormShow"

    If IsObject(Index) Then
        m_DockedForms.ItemByHandle(Index.hWnd).DockForm_Show
    Else
        m_DockedForms(Index).DockForm_Show
    End If

Exit Sub
Err_FormShow:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

Public Sub FormHide(Index As Variant)
Attribute FormHide.VB_Description = "Hides the form specified in Index"
On Error GoTo Err_FormHide
    Const constSource As String = m_constClassName & ".FormHide"

    If IsObject(Index) Then
        m_DockedForms.ItemByHandle(Index.hWnd).DockForm_Hide
    Else
        m_DockedForms(Index).DockForm_Hide
    End If

Exit Sub
Err_FormHide:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

Public Sub FormDock(Index As Variant)
Attribute FormDock.VB_Description = "Docks a form in its panel host"
On Error GoTo Err_FormDock
    Const constSource As String = m_constClassName & ".FormDock"

    If IsObject(Index) Then
        m_DockedForms.ItemByHandle(Index.hWnd).DockForm_Dock
    Else
        m_DockedForms(Index).DockForm_Dock
    End If

Exit Sub
Err_FormDock:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

Public Sub FormUndock(Index As Variant)
Attribute FormUndock.VB_Description = "Undocks a form from its panel host"
On Error GoTo Err_FormUndock
    Const constSource As String = m_constClassName & ".FormUndock"

    If IsObject(Index) Then
        m_DockedForms.ItemByHandle(Index.hWnd).DockForm_UnDock
    Else
        m_DockedForms(Index).DockForm_UnDock
    End If

Exit Sub
Err_FormUndock:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub
