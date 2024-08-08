VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton btnstart 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   3120
      Picture         =   "Formtictactoe1.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnexit_Click(Index As Integer)
End
End Sub

Private Sub btnstart_Click(Index As Integer)
Form2.Show

End Sub
