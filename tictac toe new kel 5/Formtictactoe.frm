VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0080FF80&
   Caption         =   "Form2"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14715
   LinkTopic       =   "Form2"
   ScaleHeight     =   7950
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   11760
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PLAY AGAIN"
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
      Left            =   9960
      TabIndex        =   2
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER  2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER  1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2205
      Left            =   10560
      Picture         =   "Formtictactoe.frx":0000
      Top             =   3600
      Width           =   2190
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2010
      Left            =   10560
      Picture         =   "Formtictactoe.frx":0EF3
      Top             =   840
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   8
      Left            =   6600
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   7
      Left            =   3720
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Line Line4 
      X1              =   720
      X2              =   8880
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line3 
      X1              =   720
      X2              =   9000
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   6
      Left            =   840
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   5
      Left            =   6600
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   4
      Left            =   3720
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Line Line2 
      X1              =   6120
      X2              =   6120
      Y1              =   120
      Y2              =   7560
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   3
      Left            =   6600
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   2
      Left            =   840
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   1
      Left            =   3720
      Top             =   480
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   3240
      Y1              =   120
      Y2              =   7560
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   0
      Left            =   840
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cross(8) As Boolean
Dim ball(8) As Boolean
Dim m As Integer
Dim player As Integer


Sub check_status()
If ball(0) = True And ball(1) = True And ball(2) = True Then
MsgBox "Player-1 Wins"
End If
If ball(3) = True And ball(4) = True And ball(5) = True Then
MsgBox "Player-1 Wins"
End If
If ball(6) = True And ball(7) = True And ball(8) = True Then
MsgBox "Player-1 Wins"
End If
If ball(0) = True And ball(3) = True And ball(6) = True Then
MsgBox "Player-1 Wins"
End If

If ball(1) = True And ball(4) = True And ball(7) = True Then
MsgBox "Player-1 Wins"
End If
If ball(2) = True And ball(5) = True And ball(8) = True Then
MsgBox "Player-1 Wins"
End If
If ball(0) = True And ball(4) = True And ball(8) = True Then
MsgBox "Player-1 Wins"
End If
If ball(2) = True And ball(4) = True And ball(6) = True Then
MsgBox "Player 1 Wins"
End If

If cross(0) = True And cross(1) = True And cross(2) = True Then
MsgBox "Player-2 Wins"
End If
If cross(3) = True And cross(4) = True And cross(5) = True Then
MsgBox "Player-2 Wins"
End If
If cross(6) = True And cross(7) = True And cross(8) = True Then
MsgBox "Player-2 Wins"
End If
If cross(0) = True And cross(3) = True And cross(6) = True Then
MsgBox "Player-2 Wins"
End If

If cross(1) = True And cross(4) = True And cross(7) = True Then
MsgBox "Player-2 Wins"
End If
If cross(2) = True And cross(5) = True And cross(8) = True Then
MsgBox "Player-2 Wins"
End If
If cross(0) = True And cross(4) = True And cross(8) = True Then
MsgBox "Player-2 Wins"
End If
If cross(2) = True And cross(4) = True And cross(6) = True Then
MsgBox "Player 2 wins"
End If

End Sub

Sub check_position()
For m = 0 To 8
If Image1(m).Picture = Image2.Picture Then
ball(m) = True
Else: ball(m) = False
End If
If Image1(m).Picture = Image3.Picture Then
cross(m) = True
Else
cross(m) = False
End If
Next
End Sub

Private Sub Command1_Click()
For m = 0 To 8
Image1(m).Picture = LoadPicture("")
Next
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Image1_Click(Index As Integer)
check_position
If player = 1 And cross(Index) = False And ball(Index) = False Then
Image1(Index).Picture = Image2.Picture
End If
If player = 2 And cross(Index) = False And ball(Index) = False Then
Image1(Index).Picture = Image3.Picture
End If
check_position
check_status
End Sub

Private Sub Image2_Click()
player = 1
End Sub

Private Sub Image3_Click()
player = 2
End Sub
