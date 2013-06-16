VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Chess 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RagiChess"
   ClientHeight    =   9435
   ClientLeft      =   2160
   ClientTop       =   675
   ClientWidth     =   13455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Chess Adventurer"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "first.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13455
   Begin VB.Timer Timer8 
      Left            =   10680
      Top             =   6120
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      Height          =   3015
      Left            =   0
      Picture         =   "first.frx":058A
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   100
      Top             =   0
      Width           =   3015
   End
   Begin VB.Timer Timer7 
      Interval        =   7
      Left            =   3360
      Top             =   480
   End
   Begin VB.Timer Timer6 
      Left            =   12840
      Top             =   1440
   End
   Begin VB.Timer Timer5 
      Interval        =   250
      Left            =   12840
      Top             =   600
   End
   Begin VB.Timer Timer4 
      Left            =   10440
      Top             =   1080
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Draw"
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Resign"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Left            =   10440
      Top             =   360
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option1"
      Height          =   285
      Left            =   11280
      TabIndex        =   0
      Top             =   720
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
      Height          =   285
      Left            =   11280
      TabIndex        =   69
      Top             =   1200
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   11160
      TabIndex        =   68
      Top             =   360
      Width           =   1455
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   285
         Left            =   120
         TabIndex        =   73
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "30 minutes"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   74
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Unlimited"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   72
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "15 minutes"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   71
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   70
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   240
      Top             =   8280
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Turntimer 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   135
      Left            =   13560
      TabIndex        =   102
      Top             =   240
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   238
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   10560
      TabIndex        =   101
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RagiChess"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   4320
      TabIndex        =   99
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "SET THE TIME! >>>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   98
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   97
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   96
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   95
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   94
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   93
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   92
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   91
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   90
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "      A                 B                C               D                E                F                G               H"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   89
      Top             =   8760
      Width           =   8775
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The confirmation button for the opponent only"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   10800
      TabIndex        =   88
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*Time Delay for every turn = 5 seconds"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   83
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   82
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   81
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   80
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Black"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   79
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "White"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   78
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   77
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   76
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   75
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Iskandar Setiadi and Team 2011. Hak Cipta Tim VB SMA Ricci 1. Directed by: Wuragil Darmoko"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   11880
      TabIndex        =   67
      Top             =   9120
      Width           =   9975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   66
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Turn:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   65
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   16
      Left            =   1440
      TabIndex        =   64
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   18
      Left            =   3600
      TabIndex        =   63
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   20
      Left            =   5760
      TabIndex        =   62
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   22
      Left            =   7920
      TabIndex        =   61
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   25
      Left            =   2520
      TabIndex        =   60
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   27
      Left            =   4680
      TabIndex        =   59
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   29
      Left            =   6840
      TabIndex        =   58
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   31
      Left            =   9000
      TabIndex        =   57
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   32
      Left            =   1440
      TabIndex        =   56
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   34
      Left            =   3600
      TabIndex        =   55
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   36
      Left            =   5760
      TabIndex        =   54
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   38
      Left            =   7920
      TabIndex        =   53
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   41
      Left            =   2520
      TabIndex        =   52
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   43
      Left            =   4680
      TabIndex        =   51
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   45
      Left            =   6840
      TabIndex        =   50
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   47
      Left            =   9000
      TabIndex        =   49
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   48
      Left            =   1440
      TabIndex        =   48
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   50
      Left            =   3600
      TabIndex        =   47
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   52
      Left            =   5760
      TabIndex        =   46
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   54
      Left            =   7920
      TabIndex        =   45
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   57
      Left            =   2520
      TabIndex        =   44
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   59
      Left            =   4680
      TabIndex        =   43
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   61
      Left            =   6840
      TabIndex        =   42
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   63
      Left            =   9000
      TabIndex        =   41
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   49
      Left            =   2520
      TabIndex        =   40
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   51
      Left            =   4680
      TabIndex        =   39
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   53
      Left            =   6840
      TabIndex        =   38
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   55
      Left            =   9000
      TabIndex        =   37
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   56
      Left            =   1440
      TabIndex        =   36
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   58
      Left            =   3600
      TabIndex        =   35
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   60
      Left            =   5760
      TabIndex        =   34
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   62
      Left            =   7920
      TabIndex        =   33
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   19
      Left            =   4680
      TabIndex        =   32
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   21
      Left            =   6840
      TabIndex        =   31
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   23
      Left            =   9000
      TabIndex        =   30
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   24
      Left            =   1440
      TabIndex        =   29
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   26
      Left            =   3600
      TabIndex        =   28
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   28
      Left            =   5760
      TabIndex        =   27
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   30
      Left            =   7920
      TabIndex        =   26
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   33
      Left            =   2520
      TabIndex        =   25
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   35
      Left            =   4680
      TabIndex        =   24
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   37
      Left            =   6840
      TabIndex        =   23
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   39
      Left            =   9000
      TabIndex        =   22
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   40
      Left            =   1440
      TabIndex        =   21
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   42
      Left            =   3600
      TabIndex        =   20
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   44
      Left            =   5760
      TabIndex        =   19
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   46
      Left            =   7920
      TabIndex        =   18
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   17
      Left            =   2520
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   15
      Left            =   9000
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   14
      Left            =   7920
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   13
      Left            =   6840
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   12
      Left            =   5760
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   11
      Left            =   4680
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   10
      Left            =   3600
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   2520
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   9000
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   7920
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label r 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Chess Adventurer"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnmain 
      Caption         =   "Main"
      Begin VB.Menu mnew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnhelp 
      Caption         =   "Help"
      Begin VB.Menu mnabout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Chess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public callcode, wincode, callcodetwo, whitewin, blackwin, draw As Integer
Dim moverange As String
Dim resign, x As Single
Dim f, countt, whleftcastling, whrightcastling, blleftcastling, blrightcastling, checklabel As Integer
Dim checkcountter, castlingcountter, fiftymovecountter, fiftymovecounttertwo, fiftymovecountterthree, drawstate As Integer
Dim turn, enpassant As Boolean 'menentukan hitam atau putih
Dim enpassantcallcode, enpassanttimer, timesatu, timedua, timetiga, timeempat, timedelay As Integer
Dim whitelocation, blacklocation, xvertical, xhorizontal, rint, i, movep, a, hasil, sisa As Integer
'font purposes
Dim AppPath As String
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_FONTCHANGE = &H1D
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub whitecheck()
hasil = Int(whitelocation / 8)
sisa = whitelocation Mod 8
If r(whitelocation + 7) = "o" And sisa > 0 Then
Timer8.Interval = 1000
checklabel = 0
countt = 0
End If
If r(whitelocation + 9) = "o" And sisa < 7 Then
Timer8.Interval = 1000
checklabel = 0
countt = 0
End If
'start
If whitelocation > 15 Then
If sisa < 7 Then
If r(8 * (hasil - 2) + sisa + 1) = "m" Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
End If
If sisa > 0 Then
If r(8 * (hasil - 2) + sisa - 1) = "m" Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
End If
End If
If whitelocation > 7 Then
If sisa < 6 Then
If r(8 * (hasil - 1) + sisa + 2) = "m" Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
End If
If sisa > 1 Then
If r(8 * (hasil - 1) + sisa - 2) = "m" Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
End If
End If
If whitelocation < 56 Then
If sisa < 6 Then
If r(8 * (hasil + 1) + sisa + 2) = "m" Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
End If
If sisa > 1 Then
If r(8 * (hasil + 1) + sisa - 2) = "m" Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
End If
End If
If whitelocation < 48 Then
If sisa < 7 Then
If r(8 * (hasil + 2) + sisa + 1) = "m" Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
End If
If sisa > 0 Then
If r(8 * (hasil + 2) + sisa - 1) = "m" Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
End If
End If
'ends
'peluncur start
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa - i) >= 0 Then
If r(8 * (hasil - i) + sisa - i) = "v" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * (hasil - i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa - i) >= 0 Then
If r(8 * (hasil + i) + sisa - i) = "v" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * (hasil + i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa + i) <= 7 Then
If r(8 * (hasil - i) + sisa + i) = "v" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * (hasil - i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa + i) <= 7 Then
If r(8 * (hasil + i) + sisa + i) = "v" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * (hasil + i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
'peluncur ends
'benteng start
If sisa <> 7 Then
For i = (sisa + 1) To 7
If r(8 * hasil + i) = "t" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If sisa <> 0 Then
For i = (sisa - 1) To 0 Step -1
If r(8 * hasil + i) = "t" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If hasil <> 7 Then
For i = (hasil + 1) To 7
If r(8 * i + sisa) = "t" And xvertical = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
If hasil <> 0 Then
For i = (hasil - 1) To 0 Step -1
If r(8 * i + sisa) = "t" And xvertical = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
'benteng ends
'ratu start
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa - i) >= 0 Then
If r(8 * (hasil - i) + sisa - i) = "w" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * (hasil - i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa - i) >= 0 Then
If r(8 * (hasil + i) + sisa - i) = "w" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * (hasil + i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa + i) <= 7 Then
If r(8 * (hasil - i) + sisa + i) = "w" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * (hasil - i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa + i) <= 7 Then
If r(8 * (hasil + i) + sisa + i) = "w" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * (hasil + i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
If sisa <> 7 Then
For i = (sisa + 1) To 7
If r(8 * hasil + i) = "w" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If sisa <> 0 Then
For i = (sisa - 1) To 0 Step -1
If r(8 * hasil + i) = "w" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If hasil <> 7 Then
For i = (hasil + 1) To 7
If r(8 * i + sisa) = "w" And xvertical = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
If hasil <> 0 Then
For i = (hasil - 1) To 0 Step -1
If r(8 * i + sisa) = "w" And xvertical = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 0
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
'ratu ends
End Sub
Private Sub blackcheck()
hasil = Int(blacklocation / 8)
sisa = blacklocation Mod 8
If r(blacklocation - 7) = "p" And sisa < 7 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(blacklocation - 9) = "p" And sisa > 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
'kuda start
If blacklocation > 15 Then
If sisa < 7 Then
If r(8 * (hasil - 2) + sisa + 1) = "n" Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
End If
If sisa > 0 Then
If r(8 * (hasil - 2) + sisa - 1) = "n" Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
End If
End If
If blacklocation > 7 Then
If sisa < 6 Then
If r(8 * (hasil - 1) + sisa + 2) = "n" Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
End If
If sisa > 1 Then
If r(8 * (hasil - 1) + sisa - 2) = "n" Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
End If
End If
If blacklocation < 56 Then
If sisa < 6 Then
If r(8 * (hasil + 1) + sisa + 2) = "n" Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
End If
If sisa > 1 Then
If r(8 * (hasil + 1) + sisa - 2) = "n" Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
End If
End If
If blacklocation < 48 Then
If sisa < 7 Then
If r(8 * (hasil + 2) + sisa + 1) = "n" Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
End If
If sisa > 0 Then
If r(8 * (hasil + 2) + sisa - 1) = "n" Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
End If
End If
'kuda ends
'peluncur start
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa - i) >= 0 Then
If r(8 * (hasil - i) + sisa - i) = "b" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * (hasil - i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa - i) >= 0 Then
If r(8 * (hasil + i) + sisa - i) = "b" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * (hasil + i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa + i) <= 7 Then
If r(8 * (hasil - i) + sisa + i) = "b" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * (hasil - i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa + i) <= 7 Then
If r(8 * (hasil + i) + sisa + i) = "b" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * (hasil + i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
'peluncur ends
'benteng start
If sisa <> 7 Then
For i = (sisa + 1) To 7
If r(8 * hasil + i) = "r" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If sisa <> 0 Then
For i = (sisa - 1) To 0 Step -1
If r(8 * hasil + i) = "r" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If hasil <> 7 Then
For i = (hasil + 1) To 7
If r(8 * i + sisa) = "r" And xvertical = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
If hasil <> 0 Then
For i = (hasil - 1) To 0 Step -1
If r(8 * i + sisa) = "r" And xvertical = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
'benteng ends
'ratu start
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa - i) >= 0 Then
If r(8 * (hasil - i) + sisa - i) = "q" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * (hasil - i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa - i) >= 0 Then
If r(8 * (hasil + i) + sisa - i) = "q" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * (hasil + i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa + i) <= 7 Then
If r(8 * (hasil - i) + sisa + i) = "q" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * (hasil - i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa + i) <= 7 Then
If r(8 * (hasil + i) + sisa + i) = "q" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * (hasil + i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
If sisa <> 7 Then
For i = (sisa + 1) To 7
If r(8 * hasil + i) = "q" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If sisa <> 0 Then
For i = (sisa - 1) To 0 Step -1
If r(8 * hasil + i) = "q" And xhorizontal = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If hasil <> 7 Then
For i = (hasil + 1) To 7
If r(8 * i + sisa) = "q" And xvertical = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
If hasil <> 0 Then
For i = (hasil - 1) To 0 Step -1
If r(8 * i + sisa) = "q" And xvertical = 0 Then
Timer8.Interval = 1000
countt = 0
checklabel = 1
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
'ratu ends
End Sub
Private Sub whiteturn()
For i = 0 To 63
If r(i) = "p" Or r(i) = "n" Or r(i) = "b" Or r(i) = "r" Or r(i) = "q" Or r(i) = "k" Then
r(i).Enabled = True
Else
r(i).Enabled = False
End If
Next i
End Sub
Private Sub blackturn()
For i = 0 To 63
If r(i) = "o" Or r(i) = "m" Or r(i) = "v" Or r(i) = "t" Or r(i) = "w" Or r(i) = "l" Then
r(i).Enabled = True
Else
r(i).Enabled = False
End If
Next i
End Sub
Private Sub warnaasal()
For i = 0 To 7 Step 2
r(i).BackColor = &HFFFF80
Next i
For i = 1 To 7 Step 2
r(i).BackColor = vbWhite
Next i
For i = 9 To 15 Step 2
r(i).BackColor = &HFFFF80
Next i
For i = 8 To 15 Step 2
r(i).BackColor = vbWhite
Next i
For i = 16 To 23 Step 2
r(i).BackColor = &HFFFF80
Next i
For i = 17 To 23 Step 2
r(i).BackColor = vbWhite
Next i
For i = 25 To 31 Step 2
r(i).BackColor = &HFFFF80
Next i
For i = 24 To 31 Step 2
r(i).BackColor = vbWhite
Next i
For i = 32 To 39 Step 2
r(i).BackColor = &HFFFF80
Next i
For i = 33 To 39 Step 2
r(i).BackColor = vbWhite
Next i
For i = 41 To 47 Step 2
r(i).BackColor = &HFFFF80
Next i
For i = 40 To 47 Step 2
r(i).BackColor = vbWhite
Next i
For i = 48 To 55 Step 2
r(i).BackColor = &HFFFF80
Next i
For i = 49 To 55 Step 2
r(i).BackColor = vbWhite
Next i
For i = 57 To 63 Step 2
r(i).BackColor = &HFFFF80
Next i
For i = 56 To 63 Step 2
r(i).BackColor = vbWhite
Next i
End Sub
Private Sub enbldfalse()
For i = 0 To 63
r(i).Enabled = False
Next i
End Sub
Private Sub enbldtrue()
For i = 0 To 63
r(i).Enabled = True
Next i
End Sub
Private Sub Command1_Click()
Timer3.Interval = 0
If Label2.Caption = "White" Then
MsgBox "Black Win! White has gone to hell!"
wincode = 2
Else
MsgBox "White Win! Black has gone to hell!"
wincode = 0
End If
End Sub
Private Sub Command2_Click()
Label17.Visible = True
Command2.Enabled = False
Command3.Visible = True
Command4.Visible = True
End Sub
Private Sub Command3_Click()
Timer3.Interval = 0
MsgBox "Surat damai telah ditandatangani kedua pihak!"
wincode = 1
End Sub
Private Sub Command4_Click()
Command3.Visible = False
Command4.Visible = False
Label17.Visible = False
Command2.Enabled = True
End Sub


Private Sub Form_Load()
wincode = 3
Timer6.Interval = 50
Label27.ForeColor = vbGreen
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label17.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command1.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label1.Visible = False
Label2.Visible = False
movep = 0
warnaasal
For i = 0 To 63
r(i).Visible = False
Next i
End Sub


Private Sub mnabout_Click()
About.Visible = True
End Sub
Private Sub mnew_Click()
Label29.Caption = ""
checklabel = 4
countt = 0
Picture1.Left = Picture1.Left - 4000
If Picture1.Left < -5000 Then
Picture1.Left = -5000
End If
Welcome.WindowsMediaPlayer1.URL = ""
Welcome.WindowsMediaPlayer1.Close
Welcome.Timer2.Interval = 0
Timer5.Interval = 0
Timer7.Interval = 0
Label28.Visible = False
Label27.Visible = False
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
whleftcastling = 0
whrightcastling = 0
blleftcastling = 0
blrightcastling = 0
Command2.Visible = True
Command3.Visible = False
Command4.Visible = False
Label17.Visible = False
Command1.Visible = False
resign = 0
Label16.Visible = False
timedelay = 5
If Option1.Value = True Then
timesatu = 30
timedua = 0
timetiga = 30
timeempat = 0
Timer3.Interval = 1000
End If
If Option2.Value = True Then
timesatu = 15
timedua = 0
timetiga = 15
timeempat = 0
Timer3.Interval = 1000
End If
If Option3.Value = True Then
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Else
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
End If
Frame1.Visible = False
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
For i = 0 To 63
r(i).Visible = True
Next i
turn = True
enbldtrue
warnaasal
movep = 0
Label1.Visible = True
Label2.Visible = True
For i = 16 To 47
r(i) = ""
Next i
r(0) = "r"
r(7) = "r"
r(1) = "n"
r(6) = "n"
r(2) = "b"
r(5) = "b"
r(3) = "q"
r(4) = "k"
For i = 8 To 15
r(i) = "p"
Next i
For i = 48 To 55
r(i) = "o"
Next i
r(56) = "t"
r(63) = "t"
r(57) = "m"
r(62) = "m"
r(58) = "v"
r(61) = "v"
r(59) = "w"
r(60) = "l"
Timer1.Interval = 10
End Sub
Private Sub mnexit_Click()
End
End Sub
Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
Option3.Value = False
End Sub
Private Sub Option2_Click()
Option1.Value = False
Option2.Value = True
Option3.Value = False
End Sub
Private Sub Option3_Click()
Option2.Value = False
Option1.Value = False
Option3.Value = True
End Sub



Private Sub r_Click(Index As Integer)
If callcode = Index And movep = 0 Then
warnaasal
End If
If r(callcodetwo).BackColor = vbGreen Then
warnaasal
a = 1
r(callcodetwo).Caption = r(callcode).Caption
If callcodetwo <> callcode Then
If whleftcastling = 0 And callcodetwo = 2 Then
If r(callcodetwo) = "k" Then
r(0) = ""
r(3) = "r"
r(3).Enabled = False
End If
End If
If whrightcastling = 0 And callcodetwo = 6 Then
If r(callcodetwo) = "k" Then
r(7) = ""
r(5) = "r"
r(5).Enabled = False
End If
End If
If blleftcastling = 0 And callcodetwo = 58 Then
If r(callcodetwo) = "l" Then
r(56) = ""
r(59) = "t"
r(59).Enabled = False
End If
End If
If blrightcastling = 0 And callcodetwo = 62 Then
If r(callcodetwo) = "l" Then
r(63) = ""
r(61) = "t"
r(61).Enabled = False
End If
End If
If r(callcodetwo) = "p" Or r(callcodetwo) = "o" Then
drawstate = 0
End If
r(callcode).Caption = ""
End If
End If

'Bgi 4 kasus yaitu turn true dan false, movep 0 kondisi normal, movep 1 kondisi menyerang
If movep = 0 Then
callcode = Index
End If
'Dibagian movep 0 masing-masing turn memiliki 6 cases
If turn = True Then
whiteturn
If movep = 0 Then
If r(callcode).Caption = "p" Or r(callcode).Caption = "n" Or r(callcode).Caption = "b" Or r(callcode) = "r" Or r(callcode) = "q" Or r(callcode) = "k" Then

If r(callcode) = "p" Then 'case prajurit seratus persen
If Int(callcode / 8) = 4 And enpassant = True Then
r(enpassantcallcode).BackColor = &HFFFF&
End If
If Int(callcode / 8) = 1 Then
If r(callcode + 16).Caption = "" And r(callcode + 8).Caption = "" Then
r(callcode + 16).BackColor = &HFFFF&
End If
End If
If r(callcode + 8).Caption = "" Then
r(callcode + 8).BackColor = &HFFFF&
End If
If r(callcode + 7).Caption = "o" Or r(callcode + 7).Caption = "m" Or r(callcode + 7).Caption = "v" Or r(callcode + 7) = "t" Or r(callcode + 7) = "w" Or r(callcode + 7) = "l" Then
r(callcode + 7).BackColor = &HFFFF&
End If
If r(callcode + 9).Caption = "o" Or r(callcode + 9).Caption = "m" Or r(callcode + 9).Caption = "v" Or r(callcode + 9) = "t" Or r(callcode + 9) = "w" Or r(callcode + 9) = "l" Then
r(callcode + 9).BackColor = &HFFFF&
End If
r(callcode).BackColor = &HFFFF&
End If
If r(callcode) = "n" Then  'case kuda seratus persen
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
If hasil < 7 And sisa > 1 Then
If r(8 * (hasil + 1) + sisa - 2) = "" Or r(8 * (hasil + 1) + sisa - 2) = "o" Or r(8 * (hasil + 1) + sisa - 2) = "m" Or r(8 * (hasil + 1) + sisa - 2) = "v" Or r(8 * (hasil + 1) + sisa - 2) = "t" Or r(8 * (hasil + 1) + sisa - 2) = "w" Or r(8 * (hasil + 1) + sisa - 2) = "l" Then
r(8 * (hasil + 1) + sisa - 2).BackColor = &HFFFF&
End If
End If
If hasil < 7 And sisa < 6 Then
If r(8 * (hasil + 1) + sisa + 2) = "" Or r(8 * (hasil + 1) + sisa + 2) = "o" Or r(8 * (hasil + 1) + sisa + 2) = "m" Or r(8 * (hasil + 1) + sisa + 2) = "v" Or r(8 * (hasil + 1) + sisa + 2) = "t" Or r(8 * (hasil + 1) + sisa + 2) = "w" Or r(8 * (hasil + 1) + sisa + 2) = "l" Then
r(8 * (hasil + 1) + sisa + 2).BackColor = &HFFFF&
End If
End If
If hasil > 0 And sisa > 1 Then
If r(8 * (hasil - 1) + sisa - 2) = "" Or r(8 * (hasil - 1) + sisa - 2) = "o" Or r(8 * (hasil - 1) + sisa - 2) = "m" Or r(8 * (hasil - 1) + sisa - 2) = "v" Or r(8 * (hasil - 1) + sisa - 2) = "t" Or r(8 * (hasil - 1) + sisa - 2) = "w" Or r(8 * (hasil - 1) + sisa - 2) = "l" Then
r(8 * (hasil - 1) + sisa - 2).BackColor = &HFFFF&
End If
End If
If hasil > 0 And sisa < 6 Then
If r(8 * (hasil - 1) + sisa + 2) = "" Or r(8 * (hasil - 1) + sisa + 2) = "o" Or r(8 * (hasil - 1) + sisa + 2) = "m" Or r(8 * (hasil - 1) + sisa + 2) = "v" Or r(8 * (hasil - 1) + sisa + 2) = "t" Or r(8 * (hasil - 1) + sisa + 2) = "w" Or r(8 * (hasil - 1) + sisa + 2) = "l" Then
r(8 * (hasil - 1) + sisa + 2).BackColor = &HFFFF&
End If
End If
If hasil < 6 And sisa > 0 Then
If r(8 * (hasil + 2) + sisa - 1) = "" Or r(8 * (hasil + 2) + sisa - 1) = "o" Or r(8 * (hasil + 2) + sisa - 1) = "m" Or r(8 * (hasil + 2) + sisa - 1) = "v" Or r(8 * (hasil + 2) + sisa - 1) = "t" Or r(8 * (hasil + 2) + sisa - 1) = "w" Or r(8 * (hasil + 2) + sisa - 1) = "l" Then
r(8 * (hasil + 2) + sisa - 1).BackColor = &HFFFF&
End If
End If
If hasil < 6 And sisa < 7 Then
If r(8 * (hasil + 2) + sisa + 1) = "" Or r(8 * (hasil + 2) + sisa + 1) = "o" Or r(8 * (hasil + 2) + sisa + 1) = "m" Or r(8 * (hasil + 2) + sisa + 1) = "v" Or r(8 * (hasil + 2) + sisa + 1) = "t" Or r(8 * (hasil + 2) + sisa + 1) = "w" Or r(8 * (hasil + 2) + sisa + 1) = "l" Then
r(8 * (hasil + 2) + sisa + 1).BackColor = &HFFFF&
End If
End If
If hasil > 1 And sisa > 0 Then
If r(8 * (hasil - 2) + sisa - 1) = "" Or r(8 * (hasil - 2) + sisa - 1) = "o" Or r(8 * (hasil - 2) + sisa - 1) = "m" Or r(8 * (hasil - 2) + sisa - 1) = "v" Or r(8 * (hasil - 2) + sisa - 1) = "t" Or r(8 * (hasil - 2) + sisa - 1) = "w" Or r(8 * (hasil - 2) + sisa - 1) = "l" Then
r(8 * (hasil - 2) + sisa - 1).BackColor = &HFFFF&
End If
End If
If hasil > 1 And sisa < 7 Then
If r(8 * (hasil - 2) + sisa + 1) = "" Or r(8 * (hasil - 2) + sisa + 1) = "o" Or r(8 * (hasil - 2) + sisa + 1) = "m" Or r(8 * (hasil - 2) + sisa + 1) = "v" Or r(8 * (hasil - 2) + sisa + 1) = "t" Or r(8 * (hasil - 2) + sisa + 1) = "w" Or r(8 * (hasil - 2) + sisa + 1) = "l" Then
r(8 * (hasil - 2) + sisa + 1).BackColor = &HFFFF&
End If
End If
End If

If r(callcode) = "b" Then 'case peluncur seratus persen
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa - i) >= 0 Then
If xhorizontal = 0 Then
If r(8 * (hasil - i) + sisa - i) = "" Or r(8 * (hasil - i) + sisa - i) = "o" Or r(8 * (hasil - i) + sisa - i) = "m" Or r(8 * (hasil - i) + sisa - i) = "v" Or r(8 * (hasil - i) + sisa - i) = "t" Or r(8 * (hasil - i) + sisa - i) = "w" Or r(8 * (hasil - i) + sisa - i) = "l" Then
r(8 * (hasil - i) + sisa - i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil - i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa - i) >= 0 Then
If xhorizontal = 0 Then
If r(8 * (hasil + i) + sisa - i) = "" Or r(8 * (hasil + i) + sisa - i) = "o" Or r(8 * (hasil + i) + sisa - i) = "m" Or r(8 * (hasil + i) + sisa - i) = "v" Or r(8 * (hasil + i) + sisa - i) = "t" Or r(8 * (hasil + i) + sisa - i) = "w" Or r(8 * (hasil + i) + sisa - i) = "l" Then
r(8 * (hasil + i) + sisa - i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil + i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa + i) <= 7 Then
If xhorizontal = 0 Then
If r(8 * (hasil - i) + sisa + i) = "" Or r(8 * (hasil - i) + sisa + i) = "o" Or r(8 * (hasil - i) + sisa + i) = "m" Or r(8 * (hasil - i) + sisa + i) = "v" Or r(8 * (hasil - i) + sisa + i) = "t" Or r(8 * (hasil - i) + sisa + i) = "w" Or r(8 * (hasil - i) + sisa + i) = "l" Then
r(8 * (hasil - i) + sisa + i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil - i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa + i) <= 7 Then
If xhorizontal = 0 Then
If r(8 * (hasil + i) + sisa + i) = "" Or r(8 * (hasil + i) + sisa + i) = "o" Or r(8 * (hasil + i) + sisa + i) = "m" Or r(8 * (hasil + i) + sisa + i) = "v" Or r(8 * (hasil + i) + sisa + i) = "t" Or r(8 * (hasil + i) + sisa + i) = "w" Or r(8 * (hasil + i) + sisa + i) = "l" Then
r(8 * (hasil + i) + sisa + i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil + i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
End If

If r(callcode) = "q" Then 'case ratu seratus persen
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa - i) >= 0 Then
If xhorizontal = 0 Then
If r(8 * (hasil - i) + sisa - i) = "" Or r(8 * (hasil - i) + sisa - i) = "o" Or r(8 * (hasil - i) + sisa - i) = "m" Or r(8 * (hasil - i) + sisa - i) = "v" Or r(8 * (hasil - i) + sisa - i) = "t" Or r(8 * (hasil - i) + sisa - i) = "w" Or r(8 * (hasil - i) + sisa - i) = "l" Then
r(8 * (hasil - i) + sisa - i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil - i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa - i) >= 0 Then
If xhorizontal = 0 Then
If r(8 * (hasil + i) + sisa - i) = "" Or r(8 * (hasil + i) + sisa - i) = "o" Or r(8 * (hasil + i) + sisa - i) = "m" Or r(8 * (hasil + i) + sisa - i) = "v" Or r(8 * (hasil + i) + sisa - i) = "t" Or r(8 * (hasil + i) + sisa - i) = "w" Or r(8 * (hasil + i) + sisa - i) = "l" Then
r(8 * (hasil + i) + sisa - i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil + i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa + i) <= 7 Then
If xhorizontal = 0 Then
If r(8 * (hasil - i) + sisa + i) = "" Or r(8 * (hasil - i) + sisa + i) = "o" Or r(8 * (hasil - i) + sisa + i) = "m" Or r(8 * (hasil - i) + sisa + i) = "v" Or r(8 * (hasil - i) + sisa + i) = "t" Or r(8 * (hasil - i) + sisa + i) = "w" Or r(8 * (hasil - i) + sisa + i) = "l" Then
r(8 * (hasil - i) + sisa + i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil - i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa + i) <= 7 Then
If xhorizontal = 0 Then
If r(8 * (hasil + i) + sisa + i) = "" Or r(8 * (hasil + i) + sisa + i) = "o" Or r(8 * (hasil + i) + sisa + i) = "m" Or r(8 * (hasil + i) + sisa + i) = "v" Or r(8 * (hasil + i) + sisa + i) = "t" Or r(8 * (hasil + i) + sisa + i) = "w" Or r(8 * (hasil + i) + sisa + i) = "l" Then
r(8 * (hasil + i) + sisa + i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil + i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
If sisa <> 7 Then
For i = (sisa + 1) To 7
If xhorizontal = 0 Then
If r(8 * hasil + i) = "" Or r(8 * hasil + i) = "o" Or r(8 * hasil + i) = "m" Or r(8 * hasil + i) = "v" Or r(8 * hasil + i) = "t" Or r(8 * hasil + i) = "w" Or r(8 * hasil + i) = "l" Then
r(8 * hasil + i).BackColor = &HFFFF&
End If
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If sisa <> 0 Then
For i = (sisa - 1) To 0 Step -1
If xhorizontal = 0 Then
If r(8 * hasil + i) = "" Or r(8 * hasil + i) = "o" Or r(8 * hasil + i) = "m" Or r(8 * hasil + i) = "v" Or r(8 * hasil + i) = "t" Or r(8 * hasil + i) = "w" Or r(8 * hasil + i) = "l" Then
r(8 * hasil + i).BackColor = &HFFFF&
End If
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If hasil <> 7 Then
For i = (hasil + 1) To 7
If xvertical = 0 Then
If r(8 * i + sisa) = "" Or r(8 * i + sisa) = "o" Or r(8 * i + sisa) = "m" Or r(8 * i + sisa) = "v" Or r(8 * i + sisa) = "t" Or r(8 * i + sisa) = "w" Or r(8 * i + sisa) = "l" Then
r(8 * i + sisa).BackColor = &HFFFF&
End If
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
If hasil <> 0 Then
For i = (hasil - 1) To 0 Step -1
If xvertical = 0 Then
If r(8 * i + sisa) = "" Or r(8 * i + sisa) = "o" Or r(8 * i + sisa) = "m" Or r(8 * i + sisa) = "v" Or r(8 * i + sisa) = "t" Or r(8 * i + sisa) = "w" Or r(8 * i + sisa) = "l" Then
r(8 * i + sisa).BackColor = &HFFFF&
End If
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
End If

If r(callcode) = "k" Then 'case raja seratus persen
If whleftcastling = 0 And r(1) = "" Then
If r(2) = "" And r(3) = "" Then
r(2).BackColor = &HFFFF&
End If
End If
If whrightcastling = 0 Then
If r(5) = "" And r(6) = "" Then
r(6).BackColor = &HFFFF&
End If
End If
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
If sisa > 0 Then
If r((hasil) * 8 + sisa - 1) = "" Or r((hasil) * 8 + sisa - 1) = "o" Or r((hasil) * 8 + sisa - 1) = "m" Or r((hasil) * 8 + sisa - 1) = "v" Or r((hasil) * 8 + sisa - 1) = "t" Or r((hasil) * 8 + sisa - 1) = "w" Or r((hasil) * 8 + sisa - 1) = "l" Then
r((hasil) * 8 + sisa - 1).BackColor = &HFFFF&
End If
End If
If sisa < 7 Then
If r((hasil) * 8 + sisa + 1) = "" Or r((hasil) * 8 + sisa + 1) = "o" Or r((hasil) * 8 + sisa + 1) = "m" Or r((hasil) * 8 + sisa + 1) = "v" Or r((hasil) * 8 + sisa + 1) = "t" Or r((hasil) * 8 + sisa + 1) = "w" Or r((hasil) * 8 + sisa + 1) = "l" Then
r((hasil) * 8 + sisa + 1).BackColor = &HFFFF&
End If
End If
If hasil > 0 Then
If r((hasil - 1) * 8 + sisa) = "" Or r((hasil - 1) * 8 + sisa) = "o" Or r((hasil - 1) * 8 + sisa) = "m" Or r((hasil - 1) * 8 + sisa) = "v" Or r((hasil - 1) * 8 + sisa) = "t" Or r((hasil - 1) * 8 + sisa) = "w" Or r((hasil - 1) * 8 + sisa) = "l" Then
r((hasil - 1) * 8 + sisa).BackColor = &HFFFF&
End If
End If
If sisa > 0 And hasil > 0 Then
If r((hasil - 1) * 8 + sisa - 1) = "" Or r((hasil - 1) * 8 + sisa - 1) = "o" Or r((hasil - 1) * 8 + sisa - 1) = "m" Or r((hasil - 1) * 8 + sisa - 1) = "v" Or r((hasil - 1) * 8 + sisa - 1) = "t" Or r((hasil - 1) * 8 + sisa - 1) = "w" Or r((hasil - 1) * 8 + sisa - 1) = "l" Then
r((hasil - 1) * 8 + sisa - 1).BackColor = &HFFFF&
End If
End If
If sisa < 7 And hasil > 0 Then
If r((hasil - 1) * 8 + sisa + 1) = "" Or r((hasil - 1) * 8 + sisa + 1) = "o" Or r((hasil - 1) * 8 + sisa + 1) = "m" Or r((hasil - 1) * 8 + sisa + 1) = "v" Or r((hasil - 1) * 8 + sisa + 1) = "t" Or r((hasil - 1) * 8 + sisa + 1) = "w" Or r((hasil - 1) * 8 + sisa + 1) = "l" Then
r((hasil - 1) * 8 + sisa + 1).BackColor = &HFFFF&
End If
End If
If hasil < 7 Then
If r((hasil + 1) * 8 + sisa) = "" Or r((hasil + 1) * 8 + sisa) = "o" Or r((hasil + 1) * 8 + sisa) = "m" Or r((hasil + 1) * 8 + sisa) = "v" Or r((hasil + 1) * 8 + sisa) = "t" Or r((hasil + 1) * 8 + sisa) = "w" Or r((hasil + 1) * 8 + sisa) = "l" Then
r((hasil + 1) * 8 + sisa).BackColor = &HFFFF&
End If
End If
If sisa > 0 And hasil < 7 Then
If r((hasil + 1) * 8 + sisa - 1) = "" Or r((hasil + 1) * 8 + sisa - 1) = "o" Or r((hasil + 1) * 8 + sisa - 1) = "m" Or r((hasil + 1) * 8 + sisa - 1) = "v" Or r((hasil + 1) * 8 + sisa - 1) = "t" Or r((hasil + 1) * 8 + sisa - 1) = "w" Or r((hasil + 1) * 8 + sisa - 1) = "l" Then
r((hasil + 1) * 8 + sisa - 1).BackColor = &HFFFF&
End If
End If
If sisa < 7 And hasil < 7 Then
If r((hasil + 1) * 8 + sisa + 1) = "" Or r((hasil + 1) * 8 + sisa + 1) = "o" Or r((hasil + 1) * 8 + sisa + 1) = "m" Or r((hasil + 1) * 8 + sisa + 1) = "v" Or r((hasil + 1) * 8 + sisa + 1) = "t" Or r((hasil + 1) * 8 + sisa + 1) = "w" Or r((hasil + 1) * 8 + sisa + 1) = "l" Then
r((hasil + 1) * 8 + sisa + 1).BackColor = &HFFFF&
End If
End If
End If
If r(callcode) = "r" Then  'case benteng seratus persen
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
xhorizontal = 0
xvertical = 0
If sisa <> 7 Then
For i = (sisa + 1) To 7
If xhorizontal = 0 Then
If r(8 * hasil + i) = "" Or r(8 * hasil + i) = "o" Or r(8 * hasil + i) = "m" Or r(8 * hasil + i) = "v" Or r(8 * hasil + i) = "t" Or r(8 * hasil + i) = "w" Or r(8 * hasil + i) = "l" Then
r(8 * hasil + i).BackColor = &HFFFF&
End If
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If sisa <> 0 Then
For i = (sisa - 1) To 0 Step -1
If xhorizontal = 0 Then
If r(8 * hasil + i) = "" Or r(8 * hasil + i) = "o" Or r(8 * hasil + i) = "m" Or r(8 * hasil + i) = "v" Or r(8 * hasil + i) = "t" Or r(8 * hasil + i) = "w" Or r(8 * hasil + i) = "l" Then
r(8 * hasil + i).BackColor = &HFFFF&
End If
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If hasil <> 7 Then
For i = (hasil + 1) To 7
If xvertical = 0 Then
If r(8 * i + sisa) = "" Or r(8 * i + sisa) = "o" Or r(8 * i + sisa) = "m" Or r(8 * i + sisa) = "v" Or r(8 * i + sisa) = "t" Or r(8 * i + sisa) = "w" Or r(8 * i + sisa) = "l" Then
r(8 * i + sisa).BackColor = &HFFFF&
End If
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
If hasil <> 0 Then
For i = (hasil - 1) To 0 Step -1
If xvertical = 0 Then
If r(8 * i + sisa) = "" Or r(8 * i + sisa) = "o" Or r(8 * i + sisa) = "m" Or r(8 * i + sisa) = "v" Or r(8 * i + sisa) = "t" Or r(8 * i + sisa) = "w" Or r(8 * i + sisa) = "l" Then
r(8 * i + sisa).BackColor = &HFFFF&
End If
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
End If
End If
movep = 1
Timer1.Interval = 1
Else
movep = 0
End If
End If

If turn = False Then
blackturn
If movep = 0 Then
If r(callcode).Caption = "o" Or r(callcode).Caption = "m" Or r(callcode).Caption = "v" Or r(callcode) = "t" Or r(callcode) = "w" Or r(callcode) = "l" Then
If r(callcode) = "o" Then 'case prajurit seratus persen
If Int(callcode / 8) = 3 And enpassant = True Then
r(enpassantcallcode).BackColor = &HFFFF&
End If
If Int(callcode / 8) = 6 Then
If r(callcode - 16).Caption = "" And r(callcode - 8).Caption = "" Then
r(callcode - 16).BackColor = &HFFFF&
End If
End If
If r(callcode - 8).Caption = "" Then
r(callcode - 8).BackColor = &HFFFF&
End If
If r(callcode - 7).Caption = "p" Or r(callcode - 7).Caption = "n" Or r(callcode - 7).Caption = "b" Or r(callcode - 7) = "r" Or r(callcode - 7) = "q" Or r(callcode - 7) = "k" Then
r(callcode - 7).BackColor = &HFFFF&
End If
If r(callcode - 9).Caption = "p" Or r(callcode - 9).Caption = "n" Or r(callcode - 9).Caption = "b" Or r(callcode - 9) = "r" Or r(callcode - 9) = "q" Or r(callcode - 9) = "k" Then
r(callcode - 9).BackColor = &HFFFF&
End If
r(callcode).BackColor = &HFFFF&
End If
If r(callcode) = "m" Then  'case kuda seratus persen
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
If hasil < 7 And sisa > 1 Then
If r(8 * (hasil + 1) + sisa - 2) = "" Or r(8 * (hasil + 1) + sisa - 2) = "p" Or r(8 * (hasil + 1) + sisa - 2) = "n" Or r(8 * (hasil + 1) + sisa - 2) = "b" Or r(8 * (hasil + 1) + sisa - 2) = "q" Or r(8 * (hasil + 1) + sisa - 2) = "k" Or r(8 * (hasil + 1) + sisa - 2) = "r" Then
r(8 * (hasil + 1) + sisa - 2).BackColor = &HFFFF&
End If
End If
If hasil < 7 And sisa < 6 Then
If r(8 * (hasil + 1) + sisa + 2) = "" Or r(8 * (hasil + 1) + sisa + 2) = "p" Or r(8 * (hasil + 1) + sisa + 2) = "n" Or r(8 * (hasil + 1) + sisa + 2) = "b" Or r(8 * (hasil + 1) + sisa + 2) = "q" Or r(8 * (hasil + 1) + sisa + 2) = "k" Or r(8 * (hasil + 1) + sisa + 2) = "r" Then
r(8 * (hasil + 1) + sisa + 2).BackColor = &HFFFF&
End If
End If
If hasil > 0 And sisa > 1 Then
If r(8 * (hasil - 1) + sisa - 2) = "" Or r(8 * (hasil - 1) + sisa - 2) = "p" Or r(8 * (hasil - 1) + sisa - 2) = "n" Or r(8 * (hasil - 1) + sisa - 2) = "b" Or r(8 * (hasil - 1) + sisa - 2) = "q" Or r(8 * (hasil - 1) + sisa - 2) = "k" Or r(8 * (hasil - 1) + sisa - 2) = "r" Then
r(8 * (hasil - 1) + sisa - 2).BackColor = &HFFFF&
End If
End If
If hasil > 0 And sisa < 6 Then
If r(8 * (hasil - 1) + sisa + 2) = "" Or r(8 * (hasil - 1) + sisa + 2) = "p" Or r(8 * (hasil - 1) + sisa + 2) = "n" Or r(8 * (hasil - 1) + sisa + 2) = "b" Or r(8 * (hasil - 1) + sisa + 2) = "q" Or r(8 * (hasil - 1) + sisa + 2) = "k" Or r(8 * (hasil - 1) + sisa + 2) = "r" Then
r(8 * (hasil - 1) + sisa + 2).BackColor = &HFFFF&
End If
End If
If hasil < 6 And sisa > 0 Then
If r(8 * (hasil + 2) + sisa - 1) = "" Or r(8 * (hasil + 2) + sisa - 1) = "p" Or r(8 * (hasil + 2) + sisa - 1) = "n" Or r(8 * (hasil + 2) + sisa - 1) = "b" Or r(8 * (hasil + 2) + sisa - 1) = "q" Or r(8 * (hasil + 2) + sisa - 1) = "k" Or r(8 * (hasil + 2) + sisa - 1) = "r" Then
r(8 * (hasil + 2) + sisa - 1).BackColor = &HFFFF&
End If
End If
If hasil < 6 And sisa < 7 Then
If r(8 * (hasil + 2) + sisa + 1) = "" Or r(8 * (hasil + 2) + sisa + 1) = "p" Or r(8 * (hasil + 2) + sisa + 1) = "n" Or r(8 * (hasil + 2) + sisa + 1) = "b" Or r(8 * (hasil + 2) + sisa + 1) = "q" Or r(8 * (hasil + 2) + sisa + 1) = "k" Or r(8 * (hasil + 2) + sisa + 1) = "r" Then
r(8 * (hasil + 2) + sisa + 1).BackColor = &HFFFF&
End If
End If
If hasil > 1 And sisa > 0 Then
If r(8 * (hasil - 2) + sisa - 1) = "" Or r(8 * (hasil - 2) + sisa - 1) = "p" Or r(8 * (hasil - 2) + sisa - 1) = "n" Or r(8 * (hasil - 2) + sisa - 1) = "b" Or r(8 * (hasil - 2) + sisa - 1) = "q" Or r(8 * (hasil - 2) + sisa - 1) = "k" Or r(8 * (hasil - 2) + sisa - 1) = "r" Then
r(8 * (hasil - 2) + sisa - 1).BackColor = &HFFFF&
End If
End If
If hasil > 1 And sisa < 7 Then
If r(8 * (hasil - 2) + sisa + 1) = "" Or r(8 * (hasil - 2) + sisa + 1) = "p" Or r(8 * (hasil - 2) + sisa + 1) = "n" Or r(8 * (hasil - 2) + sisa + 1) = "b" Or r(8 * (hasil - 2) + sisa + 1) = "q" Or r(8 * (hasil - 2) + sisa + 1) = "k" Or r(8 * (hasil - 2) + sisa + 1) = "r" Then
r(8 * (hasil - 2) + sisa + 1).BackColor = &HFFFF&
End If
End If
End If
If r(callcode) = "v" Then  'case peluncur seratus persen
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa - i) >= 0 Then
If xhorizontal = 0 Then
If r(8 * (hasil - i) + sisa - i) = "" Or r(8 * (hasil - i) + sisa - i) = "p" Or r(8 * (hasil - i) + sisa - i) = "n" Or r(8 * (hasil - i) + sisa - i) = "b" Or r(8 * (hasil - i) + sisa - i) = "r" Or r(8 * (hasil - i) + sisa - i) = "q" Or r(8 * (hasil - i) + sisa - i) = "k" Then
r(8 * (hasil - i) + sisa - i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil - i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa - i) >= 0 Then
If xhorizontal = 0 Then
If r(8 * (hasil + i) + sisa - i) = "" Or r(8 * (hasil + i) + sisa - i) = "p" Or r(8 * (hasil + i) + sisa - i) = "n" Or r(8 * (hasil + i) + sisa - i) = "b" Or r(8 * (hasil + i) + sisa - i) = "r" Or r(8 * (hasil + i) + sisa - i) = "q" Or r(8 * (hasil + i) + sisa - i) = "k" Then
r(8 * (hasil + i) + sisa - i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil + i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa + i) <= 7 Then
If xhorizontal = 0 Then
If r(8 * (hasil - i) + sisa + i) = "" Or r(8 * (hasil - i) + sisa + i) = "p" Or r(8 * (hasil - i) + sisa + i) = "n" Or r(8 * (hasil - i) + sisa + i) = "b" Or r(8 * (hasil - i) + sisa + i) = "r" Or r(8 * (hasil - i) + sisa + i) = "q" Or r(8 * (hasil - i) + sisa + i) = "k" Then
r(8 * (hasil - i) + sisa + i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil - i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa + i) <= 7 Then
If xhorizontal = 0 Then
If r(8 * (hasil + i) + sisa + i) = "" Or r(8 * (hasil + i) + sisa + i) = "p" Or r(8 * (hasil + i) + sisa + i) = "n" Or r(8 * (hasil + i) + sisa + i) = "b" Or r(8 * (hasil + i) + sisa + i) = "r" Or r(8 * (hasil + i) + sisa + i) = "q" Or r(8 * (hasil + i) + sisa + i) = "k" Then
r(8 * (hasil + i) + sisa + i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil + i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
End If
If r(callcode) = "w" Then 'case ratu seratus persen
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa - i) >= 0 Then
If xhorizontal = 0 Then
If r(8 * (hasil - i) + sisa - i) = "" Or r(8 * (hasil - i) + sisa - i) = "p" Or r(8 * (hasil - i) + sisa - i) = "n" Or r(8 * (hasil - i) + sisa - i) = "b" Or r(8 * (hasil - i) + sisa - i) = "r" Or r(8 * (hasil - i) + sisa - i) = "q" Or r(8 * (hasil - i) + sisa - i) = "k" Then
r(8 * (hasil - i) + sisa - i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil - i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa - i) >= 0 Then
If xhorizontal = 0 Then
If r(8 * (hasil + i) + sisa - i) = "" Or r(8 * (hasil + i) + sisa - i) = "p" Or r(8 * (hasil + i) + sisa - i) = "n" Or r(8 * (hasil + i) + sisa - i) = "b" Or r(8 * (hasil + i) + sisa - i) = "r" Or r(8 * (hasil + i) + sisa - i) = "q" Or r(8 * (hasil + i) + sisa - i) = "k" Then
r(8 * (hasil + i) + sisa - i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil + i) + sisa - i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil - i) >= 0 And (sisa + i) <= 7 Then
If xhorizontal = 0 Then
If r(8 * (hasil - i) + sisa + i) = "" Or r(8 * (hasil - i) + sisa + i) = "p" Or r(8 * (hasil - i) + sisa + i) = "n" Or r(8 * (hasil - i) + sisa + i) = "b" Or r(8 * (hasil - i) + sisa + i) = "r" Or r(8 * (hasil - i) + sisa + i) = "q" Or r(8 * (hasil - i) + sisa + i) = "k" Then
r(8 * (hasil - i) + sisa + i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil - i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
For i = 1 To 7
If (hasil + i) <= 7 And (sisa + i) <= 7 Then
If xhorizontal = 0 Then
If r(8 * (hasil + i) + sisa + i) = "" Or r(8 * (hasil + i) + sisa + i) = "p" Or r(8 * (hasil + i) + sisa + i) = "n" Or r(8 * (hasil + i) + sisa + i) = "b" Or r(8 * (hasil + i) + sisa + i) = "r" Or r(8 * (hasil + i) + sisa + i) = "q" Or r(8 * (hasil + i) + sisa + i) = "k" Then
r(8 * (hasil + i) + sisa + i).BackColor = &HFFFF&
End If
End If
If r(8 * (hasil + i) + sisa + i) <> "" Then
xhorizontal = 1
End If
End If
Next i
xhorizontal = 0
If sisa <> 7 Then
For i = (sisa + 1) To 7
If xhorizontal = 0 Then
If r(8 * hasil + i) = "" Or r(8 * hasil + i) = "p" Or r(8 * hasil + i) = "n" Or r(8 * hasil + i) = "b" Or r(8 * hasil + i) = "r" Or r(8 * hasil + i) = "q" Or r(8 * hasil + i) = "k" Then
r(8 * hasil + i).BackColor = &HFFFF&
End If
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If sisa <> 0 Then
For i = (sisa - 1) To 0 Step -1
If xhorizontal = 0 Then
If r(8 * hasil + i) = "" Or r(8 * hasil + i) = "p" Or r(8 * hasil + i) = "n" Or r(8 * hasil + i) = "b" Or r(8 * hasil + i) = "r" Or r(8 * hasil + i) = "q" Or r(8 * hasil + i) = "k" Then
r(8 * hasil + i).BackColor = &HFFFF&
End If
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If hasil <> 7 Then
For i = (hasil + 1) To 7
If xvertical = 0 Then
If r(8 * i + sisa) = "" Or r(8 * i + sisa) = "p" Or r(8 * i + sisa) = "n" Or r(8 * i + sisa) = "b" Or r(8 * i + sisa) = "r" Or r(8 * i + sisa) = "q" Or r(8 * i + sisa) = "k" Then
r(8 * i + sisa).BackColor = &HFFFF&
End If
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
If hasil <> 0 Then
For i = (hasil - 1) To 0 Step -1
If xvertical = 0 Then
If r(8 * i + sisa) = "" Or r(8 * i + sisa) = "p" Or r(8 * i + sisa) = "n" Or r(8 * i + sisa) = "b" Or r(8 * i + sisa) = "r" Or r(8 * i + sisa) = "q" Or r(8 * i + sisa) = "k" Then
r(8 * i + sisa).BackColor = &HFFFF&
End If
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
End If

If r(callcode) = "l" Then 'case raja seratus persen
If blleftcastling = 0 And r(57) = "" Then
If r(58) = "" And r(59) = "" Then
r(58).BackColor = &HFFFF&
End If
End If
If blrightcastling = 0 Then
If r(61) = "" And r(62) = "" Then
r(62).BackColor = &HFFFF&
End If
End If
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
If sisa > 0 Then
If r((hasil) * 8 + sisa - 1) = "" Or r((hasil) * 8 + sisa - 1) = "p" Or r((hasil) * 8 + sisa - 1) = "n" Or r((hasil) * 8 + sisa - 1) = "b" Or r((hasil) * 8 + sisa - 1) = "q" Or r((hasil) * 8 + sisa - 1) = "r" Or r((hasil) * 8 + sisa - 1) = "k" Then
r((hasil) * 8 + sisa - 1).BackColor = &HFFFF&
End If
End If
If sisa < 7 Then
If r((hasil) * 8 + sisa + 1) = "" Or r((hasil) * 8 + sisa + 1) = "p" Or r((hasil) * 8 + sisa + 1) = "n" Or r((hasil) * 8 + sisa + 1) = "b" Or r((hasil) * 8 + sisa + 1) = "q" Or r((hasil) * 8 + sisa + 1) = "r" Or r((hasil) * 8 + sisa + 1) = "k" Then
r((hasil) * 8 + sisa + 1).BackColor = &HFFFF&
End If
End If
If hasil > 0 Then
If r((hasil - 1) * 8 + sisa) = "" Or r((hasil - 1) * 8 + sisa) = "p" Or r((hasil - 1) * 8 + sisa) = "n" Or r((hasil - 1) * 8 + sisa) = "b" Or r((hasil - 1) * 8 + sisa) = "q" Or r((hasil - 1) * 8 + sisa) = "r" Or r((hasil - 1) * 8 + sisa) = "k" Then
r((hasil - 1) * 8 + sisa).BackColor = &HFFFF&
End If
End If
If sisa > 0 And hasil > 0 Then
If r((hasil - 1) * 8 + sisa - 1) = "" Or r((hasil - 1) * 8 + sisa - 1) = "p" Or r((hasil - 1) * 8 + sisa - 1) = "n" Or r((hasil - 1) * 8 + sisa - 1) = "b" Or r((hasil - 1) * 8 + sisa - 1) = "q" Or r((hasil - 1) * 8 + sisa - 1) = "r" Or r((hasil - 1) * 8 + sisa - 1) = "k" Then
r((hasil - 1) * 8 + sisa - 1).BackColor = &HFFFF&
End If
End If
If sisa < 7 And hasil > 0 Then
If r((hasil - 1) * 8 + sisa + 1) = "" Or r((hasil - 1) * 8 + sisa + 1) = "p" Or r((hasil - 1) * 8 + sisa + 1) = "n" Or r((hasil - 1) * 8 + sisa + 1) = "b" Or r((hasil - 1) * 8 + sisa + 1) = "q" Or r((hasil - 1) * 8 + sisa + 1) = "r" Or r((hasil - 1) * 8 + sisa + 1) = "k" Then
r((hasil - 1) * 8 + sisa + 1).BackColor = &HFFFF&
End If
End If
If hasil < 7 Then
If r((hasil + 1) * 8 + sisa) = "" Or r((hasil + 1) * 8 + sisa) = "p" Or r((hasil + 1) * 8 + sisa) = "n" Or r((hasil + 1) * 8 + sisa) = "b" Or r((hasil + 1) * 8 + sisa) = "q" Or r((hasil + 1) * 8 + sisa) = "r" Or r((hasil + 1) * 8 + sisa) = "k" Then
r((hasil + 1) * 8 + sisa).BackColor = &HFFFF&
End If
End If
If sisa > 0 And hasil < 7 Then
If r((hasil + 1) * 8 + sisa - 1) = "" Or r((hasil + 1) * 8 + sisa - 1) = "p" Or r((hasil + 1) * 8 + sisa - 1) = "n" Or r((hasil + 1) * 8 + sisa - 1) = "b" Or r((hasil + 1) * 8 + sisa - 1) = "q" Or r((hasil + 1) * 8 + sisa - 1) = "r" Or r((hasil + 1) * 8 + sisa - 1) = "k" Then
r((hasil + 1) * 8 + sisa - 1).BackColor = &HFFFF&
End If
End If
If sisa < 7 And hasil < 7 Then
If r((hasil + 1) * 8 + sisa + 1) = "" Or r((hasil + 1) * 8 + sisa + 1) = "p" Or r((hasil + 1) * 8 + sisa + 1) = "n" Or r((hasil + 1) * 8 + sisa + 1) = "b" Or r((hasil + 1) * 8 + sisa + 1) = "q" Or r((hasil + 1) * 8 + sisa + 1) = "r" Or r((hasil + 1) * 8 + sisa + 1) = "k" Then
r((hasil + 1) * 8 + sisa + 1).BackColor = &HFFFF&
End If
End If
End If
If r(callcode) = "t" Then  'case benteng seratus persen
r(callcode).BackColor = &HFFFF&
hasil = Int(callcode / 8)
sisa = callcode Mod 8
If sisa <> 7 Then
For i = (sisa + 1) To 7
If xhorizontal = 0 Then
If r(8 * hasil + i) = "" Or r(8 * hasil + i) = "p" Or r(8 * hasil + i) = "n" Or r(8 * hasil + i) = "b" Or r(8 * hasil + i) = "q" Or r(8 * hasil + i) = "k" Or r(8 * hasil + i) = "r" Then
r(8 * hasil + i).BackColor = &HFFFF&
End If
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If sisa <> 0 Then
For i = (sisa - 1) To 0 Step -1
If xhorizontal = 0 Then
If r(8 * hasil + i) = "" Or r(8 * hasil + i) = "p" Or r(8 * hasil + i) = "n" Or r(8 * hasil + i) = "b" Or r(8 * hasil + i) = "q" Or r(8 * hasil + i) = "k" Or r(8 * hasil + i) = "r" Then
r(8 * hasil + i).BackColor = &HFFFF&
End If
End If
If r(8 * hasil + i) <> "" Then
xhorizontal = 1
End If
Next i
xhorizontal = 0
End If
If hasil <> 7 Then
For i = (hasil + 1) To 7
If xvertical = 0 Then
If r(8 * i + sisa) = "" Or r(8 * i + sisa) = "p" Or r(8 * i + sisa) = "n" Or r(8 * i + sisa) = "b" Or r(8 * i + sisa) = "q" Or r(8 * i + sisa) = "k" Or r(8 * i + sisa) = "r" Then
r(8 * i + sisa).BackColor = &HFFFF&
End If
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
If hasil <> 0 Then
For i = (hasil - 1) To 0 Step -1
If xvertical = 0 Then
If r(8 * i + sisa) = "" Or r(8 * i + sisa) = "p" Or r(8 * i + sisa) = "n" Or r(8 * i + sisa) = "b" Or r(8 * i + sisa) = "q" Or r(8 * i + sisa) = "k" Or r(8 * i + sisa) = "r" Then
r(8 * i + sisa).BackColor = &HFFFF&
End If
End If
If r(8 * i + sisa) <> "" Then
xvertical = 1
End If
Next i
xvertical = 0
End If
End If
End If
movep = 1
Timer1.Interval = 1
Else
movep = 0
End If
End If
If callcode <> Index And movep = 0 Then
If a = 1 Then
resign = resign + 1
timedelay = 5
If enpassantcallcode = callcodetwo And r(callcodetwo).Caption = "o" Then
r(callcodetwo + 8) = ""
Timer8.Interval = 1000
checklabel = 2
countt = 0
End If
If enpassantcallcode = callcodetwo And r(callcodetwo).Caption = "p" Then
r(callcodetwo - 8) = ""
Timer8.Interval = 1000
checklabel = 2
countt = 0
End If
enpassanttimer = enpassanttimer - 1
If enpassanttimer = 0 Then
enpassant = False
End If
Select Case turn
Case True
turn = False
Case False
turn = True
End Select
Timer1.Interval = 1
a = 0
End If
End If
End Sub
Private Sub r_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
callcodetwo = Index
For i = 0 To 63
If callcodetwo = i And r(i).BackColor = &HFFFF& Then
r(i).BackColor = vbGreen
End If
If callcodetwo <> i And r(i).BackColor = vbGreen Then
r(i).BackColor = &HFFFF&
End If
Next i
End Sub
Private Sub Timer1_Timer()
For i = 0 To 63
If r(i).Caption = "k" Then
whitelocation = i
End If
If r(i).Caption = "l" Then
blacklocation = i
End If
Next i
If turn = True And r(callcode) <> "k" Then
whitecheck
End If
If turn = False And r(callcode) <> "l" Then
blackcheck
End If
fiftymovecountter = 0
If callcodetwo <> callcode Then
For i = 0 To 63
If r(i).Caption <> "" Then
fiftymovecountter = fiftymovecountter + 1
End If
Next i
fiftymovecountterthree = fiftymovecounttertwo
fiftymovecounttertwo = fiftymovecountter
If fiftymovecountterthree = fiftymovecounttertwo Then
drawstate = drawstate + 1
Else
drawstate = 0
End If
If drawstate = 51 Then
Timer3.Interval = 0
MsgBox "D-R-A-W!"
wincode = 1
End If
End If
If resign > 29 Then
Command1.Visible = True
End If
If movep = 1 Then
For i = 0 To 63
r(i).Enabled = False
If r(i).BackColor = vbGreen Or r(i).BackColor = &HFFFF& Then
r(i).Enabled = True
End If
Next i
End If
Timer4.Interval = 1
Timer1.Interval = 0
End Sub
Private Sub Timer2_Timer()
If Label3.Left > -10000 Then
Label3.Left = Label3.Left - 100
Else
Label3.Left = 12000
End If
End Sub
Private Sub Timer3_Timer()
If timedelay <> 0 Then
timedelay = timedelay - 1
Else
If Label2.Caption = "White" Then
If timesatu = 0 And timedua = 0 Then
MsgBox " Black Win, and White lose!"
Timer3.Interval = 0
timesatu = 0
timedua = 0
wincode = 2
End If
If timedua > 0 Then
timedua = timedua - 1
Else
timedua = 59
timesatu = timesatu - 1
End If
Else
If timetiga = 0 And timeempat = 0 Then
MsgBox " White Win, and Black lose!"
Timer3.Interval = 0
timetiga = 0
timeempat = 0
wincode = 0
End If
If timeempat > 0 Then
timeempat = timeempat - 1
Else
timeempat = 59
timetiga = timetiga - 1
End If
End If
End If
Label8.Caption = timesatu
Label10.Caption = timedua
If timedua = 0 Then
Label10.Caption = "00"
End If
Label13.Caption = timetiga
Label15.Caption = timeempat
If timeempat = 0 Then
Label15.Caption = "00"
End If
End Sub

Private Sub Timer4_Timer()
For i = 0 To 7
If r(i) = "o" Then
Promotion.Visible = True
End If
Next i
For i = 56 To 63
If r(i) = "p" Then
Promotion.Visible = True
End If
Next i
If movep = 0 Then
If callcodetwo - callcode = 16 And r(callcodetwo) = "p" Then
enpassant = True
enpassanttimer = 3
enpassantcallcode = callcode + 8
End If
If callcode - callcodetwo = 16 And r(callcodetwo) = "o" Then
enpassant = True
enpassanttimer = 2
enpassantcallcode = callcode - 8
End If
If turn = True Then
whiteturn
End If
If turn = False Then
blackturn
End If
End If
If r(0) <> "r" Or r(4) <> "k" Then
whleftcastling = 1
End If
If r(4) <> "k" Or r(7) <> "r" Then
whrightcastling = 1
End If
If r(56) <> "t" Or r(60) <> "l" Then
blleftcastling = 1
End If
If r(60) <> "l" Or r(63) <> "t" Then
blrightcastling = 1
End If
whitewin = 0
blackwin = 0
For i = 0 To 63
If r(i).Caption = "k" Then
whitewin = 1
End If
If r(i).Caption = "l" Then
blackwin = 1
End If
Next i
If whitewin = 0 Then
Promotion.Visible = False
MsgBox "Black has captured the king of white!"
wincode = 2
End If
If blackwin = 0 Then
Promotion.Visible = False
MsgBox "White has captured the king of black!"
wincode = 0
End If
Timer4.Interval = 0
End Sub

Private Sub Timer5_Timer()
If Label27.ForeColor = vbGreen Then
Label27.ForeColor = vbRed
Else
Label27.ForeColor = vbGreen
End If
If Picture1.Visible = True Then
Picture1.Visible = False
Else
Picture1.Visible = True
End If
End Sub

Private Sub Timer6_Timer()
If wincode <> 3 Then
Result.Visible = True
Chess.Enabled = False
End If
End Sub

Private Sub Timer7_Timer()
x = x + 0.1
Label28.Top = Label28.Top + Sin(x) * 200 - Cos(x) * 200
Label28.Left = Label28.Left - Sin(x) * 200 + Cos(x) * 200
If x = 360 Then
x = 0
End If
End Sub

Private Sub Timer8_Timer()
f = f + 1
If checklabel = 0 And countt = 0 Then
If turn = True Then
Label29.Caption = "Check @ White!"
End If
If f = 1 Then
WindowsMediaPlayer1.URL = "D:\Chessmaster Grand Finale Project\06.CRYSTAL.mp3"
End If
End If
If checklabel = 1 And countt = 0 Then
If turn = False Then
Label29.Caption = "Check @ Black!"
End If
If f = 1 Then
WindowsMediaPlayer1.URL = "D:\Chessmaster Grand Finale Project\06.CRYSTAL.mp3"
End If
End If
If checklabel = 2 And countt = 0 Then
Label29.Caption = "En Passant"
End If
If f = 5 Then
WindowsMediaPlayer1.URL = ""
Label29.Caption = ""
Timer8.Interval = 0
f = 0
countt = 1
End If
End Sub

Private Sub Turntimer_Timer()
If turn = True Then
Label2.Caption = "White"
Else
Label2.Caption = "Black"
End If
End Sub
'Iskandar Setiadi Collection XII.IPA 2010/2011

