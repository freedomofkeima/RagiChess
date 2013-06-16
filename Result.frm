VERSION 5.00
Begin VB.Form Result 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RagiChess"
   ClientHeight    =   2655
   ClientLeft      =   5955
   ClientTop       =   3075
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4560
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim menit, detik As Integer
Private Sub Command1_Click()
MsgBox "Thanks for Playing RagiChess!"
Welcome.WindowsMediaPlayer1.URL = ""
Welcome.WindowsMediaPlayer1.Close
End
End Sub

Private Sub Form_Load()
Chess.WindowsMediaPlayer1.Close
Welcome.WindowsMediaPlayer1.URL = "D:\Chessmaster Grand Finale Project\02 - The First Noel.mp3"
Label1.ForeColor = vbWhite
Timer2.Interval = 1000
menit = 3
detik = 9
If Chess.wincode = 0 Then
Label1.Caption = "White wins!"
End If
If Chess.wincode = 1 Then
Label1.Caption = "Draw!"
End If
If Chess.wincode = 2 Then
Label1.Caption = "Black wins!"
End If
End Sub

Private Sub Timer1_Timer()
If Label1.ForeColor = vbWhite Then
Label1.ForeColor = vbGreen
Else
Label1.ForeColor = vbWhite
End If
If Label1.Top < 1200 Then
Label1.Top = Label1.Top + 10
End If
End Sub

Private Sub Timer2_Timer()
If detik > 0 Then
detik = detik - 1
End If
If detik = 0 And menit = 0 Then
Welcome.WindowsMediaPlayer1.URL = ""
Welcome.WindowsMediaPlayer1.Close
MsgBox "Thanks for Playing RagiChess!"
End
End If
If detik = 0 Then
menit = menit - 1
detik = 60
End If
End Sub
