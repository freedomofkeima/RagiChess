VERSION 5.00
Begin VB.Form Promotion 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promote Pawns"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3705
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000002&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000002&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000002&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000002&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "bv"
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rt"
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "nm"
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "qw"
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bishop"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rook"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Knight"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Queen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Promotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
If Int(Chess.callcodetwo / 8) > 5 Then
For i = 56 To 63
If (Chess.r(i)) = "p" Then
Chess.r(i).Caption = "q"
End If
Next i
End If
If Int(Chess.callcodetwo / 8) < 2 Then
For i = 0 To 7
If (Chess.r(i)) = "o" Then
Chess.r(i).Caption = "w"
End If
Next i
End If
End If
If Option2.Value = True Then
If Int(Chess.callcodetwo / 8) > 5 Then
For i = 56 To 63
If (Chess.r(i)) = "p" Then
Chess.r(i).Caption = "n"
End If
Next i
End If
If Int(Chess.callcodetwo / 8) < 2 Then
For i = 0 To 7
If (Chess.r(i)) = "o" Then
Chess.r(i).Caption = "m"
End If
Next i
End If
End If
If Option3.Value = True Then
If Int(Chess.callcodetwo / 8) > 5 Then
For i = 56 To 63
If (Chess.r(i)) = "p" Then
Chess.r(i).Caption = "r"
End If
Next i
End If
If Int(Chess.callcodetwo / 8) < 2 Then
For i = 0 To 7
If (Chess.r(i)) = "o" Then
Chess.r(i).Caption = "t"
End If
Next i
End If
End If
If Option4.Value = True Then
If Int(Chess.callcodetwo / 8) > 5 Then
For i = 56 To 63
If (Chess.r(i)) = "p" Then
Chess.r(i).Caption = "b"
End If
Next i
End If
If Int(Chess.callcodetwo / 8) < 2 Then
For i = 0 To 7
If (Chess.r(i)) = "o" Then
Chess.r(i).Caption = "v"
End If
Next i
End If
End If
Promotion.Visible = False
Chess.Enabled = True
Chess.Visible = True
End Sub

Private Sub Form_Load()
Option1.Value = True
Option2.Value = False
Option3.Value = False
Option4.Value = False
Chess.Enabled = False
End Sub

Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
Option3.Value = False
Option4.Value = False
End Sub

Private Sub Option2_Click()
Option1.Value = False
Option2.Value = True
Option3.Value = False
Option4.Value = False
End Sub

Private Sub Option3_Click()
Option2.Value = False
Option3.Value = True
Option1.Value = False
Option4.Value = False
End Sub

Private Sub Option4_Click()
Option2.Value = False
Option4.Value = True
Option3.Value = False
Option1.Value = False
End Sub
