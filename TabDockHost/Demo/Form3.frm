VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form3"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   LinkTopic       =   "Form3"
   ScaleHeight     =   2610
   ScaleWidth      =   3045
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Form3.frx":0000
      Left            =   240
      List            =   "Form3.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   12
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form3.frx":001B
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1092
      Left            =   60
      TabIndex        =   4
      Top             =   1440
      Width           =   2835
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Type a text here please"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "True or False?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   2835
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "FORM3:MOUSEDOWN"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "FORM3:MOUSEUP"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Label3.Move 40, Label3.Top, Me.ScaleWidth - 80, Me.ScaleHeight - (Label3.Top + 20)
End Sub
'-- end code
