VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5948
      _ExtentY        =   4043
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form5.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfText.RightMargin = 10000
End Sub
'-- end code
