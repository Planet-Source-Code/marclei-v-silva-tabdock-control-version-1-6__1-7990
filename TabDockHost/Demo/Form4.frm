VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1536
   ClientLeft      =   0
   ClientTop       =   -48
   ClientWidth     =   3756
   LinkTopic       =   "Form4"
   ScaleHeight     =   1536
   ScaleWidth      =   3756
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   552
      Left            =   300
      ScaleHeight     =   552
      ScaleWidth      =   2952
      TabIndex        =   1
      Top             =   660
      Width           =   2952
      Begin VB.Label Label1 
         Caption         =   $"Form4.frx":0000
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3252
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   1332
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   2350
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Readme"
            Object.ToolTipText     =   "Show TabDock Control ReadMe file"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comments"
            Object.ToolTipText     =   "Show user comments about the control"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Revisions"
            Object.ToolTipText     =   "Revisions made to this control"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MDIForm"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    ' resize controls
    TabStrip1.Move 30, 20, Me.ScaleWidth - 60, Me.ScaleHeight - 40
    Picture1.Move TabStrip1.Left + 20, _
                TabStrip1.Top + 350, _
                TabStrip1.Width - 50, _
                TabStrip1.Height - 400
    Label1.Move 50, 50, Picture1.ScaleWidth - 100, Picture1.ScaleHeight - 100
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Index
        Case 1
            ' load readme file
            MDIForm1.LoadNewDoc App.Path & "\readme.rtf"
        Case 2
            ' load comments file
            MDIForm1.LoadNewDoc App.Path & "\comments.rtf"
        Case 3
            ' load revisions file
            MDIForm1.LoadNewDoc App.Path & "\revisions.rtf"
        Case 4
            ' load revisions file
            MDIForm1.LoadNewDoc App.Path & "\project.rtf"
    End Select
    On Error Resume Next
    Me.SetFocus
End Sub
'-- end code
