VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form2"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   1800
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0000
            Key             =   "closed"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2412
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   4255
      _Version        =   393217
      Indentation     =   265
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With TreeView1.Nodes
        .Add , , , "Item 1", "closed", "closed"
        .Add 1, tvwChild, , "SubItem 1", "closed", "closed"
        .Add , , , "Item 2", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 1", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 2", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 3", "closed", "closed"
    End With
End Sub

Private Sub Form_Resize()
    TreeView1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
'-- end code
