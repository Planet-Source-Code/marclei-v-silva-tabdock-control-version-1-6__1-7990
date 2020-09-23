VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   2508
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   3744
   LinkTopic       =   "Form6"
   ScaleHeight     =   2508
   ScaleWidth      =   3744
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1572
      Left            =   240
      ScaleHeight     =   1572
      ScaleWidth      =   3012
      TabIndex        =   1
      Top             =   480
      Width           =   3012
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "Form6.frx":0000
         Top             =   120
         Width           =   2412
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2172
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   3831
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alphabetic"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Categorized"
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
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Text1.Text = Replace(Text1.Text, vbCrLf, Chr(32))
    Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    TabStrip1.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - 20
    Picture1.Move TabStrip1.Left + 20, _
                TabStrip1.Top + 300, _
                TabStrip1.Width - 50, _
                TabStrip1.Height - 350
    Text1.Move 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub
'-- end code
