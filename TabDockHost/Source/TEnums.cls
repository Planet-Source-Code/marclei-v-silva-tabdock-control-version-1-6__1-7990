VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************************************
' Project:           TabDock
' Module:            Enums
' Description:       Several Constant definitions for this project
' Created by:        Marclei V Silva
' Machine:           ZEUS
' Date-Time:         12/05/20000:43:08
' Revisions:
' ********************************************************************Option Explicit

Public Enum tdMenuFlags
    tdmUnChecked = &H0&
    tdmChecked = &H8&
    tdmDisabled = &H2&
    tdmGrayed = &H1&
    tdmSeparator = &H800&
    tdmString = &H0&
End Enum

Public Enum tdBorderStyles
    bdrNone = 0
    bdrRaisedOuter = 1
    bdrRaisedInner = 2
    bdrRaised = 3
    bdrSunkenOuter = 4
    bdrSunkenInner = 5
    bdrSunken = 6
    bdrEtched = 7
    bdrBump = 8
    bdrMono = 9
    bdrFlat = 10
    bdrSoft = 11
End Enum

' Docking States
Public Enum tdDockedState
    tdDocked = 1
    tdUndocked = 2
End Enum

' align properties for each panel that is created
' for the docking engine
Public Enum tdAlignProperty
    tdAlignNone = 0     ' Floating host not implemented
    tdAlignTop = 1      ' Top host
    tdAlignBottom = 2   ' Bottom host
    tdAlignLeft = 3     ' Left Host
    tdAlignRight = 4    ' Right Host
End Enum

' styles for the docked form
Public Enum tdDockStyles
    tdDockLeft = &H2        ' Allow dock left
    tdDockRight = &H4       ' Allow dock right
    tdDockBottom = &H8      ' Allow dock bottom
    tdDockTop = &H10        ' Allow dock top
    tdDockFloat = &H20      ' Allow floating
    tdShowInvisible = &H40  ' Visibility at Show() method
End Enum
'-- end code
