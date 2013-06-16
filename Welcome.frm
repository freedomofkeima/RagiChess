VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Welcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                                                                 ~::Welcome  to  RagiChess::~"
   ClientHeight    =   5355
   ClientLeft      =   2790
   ClientTop       =   2655
   ClientWidth     =   11745
   ControlBox      =   0   'False
   DrawMode        =   12  'Nop
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Welcome.frx":0000
   ScaleHeight     =   5355
   ScaleWidth      =   11745
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   11160
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   4920
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   12120
      TabIndex        =   4
      Top             =   3840
      Width           =   495
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
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   5040
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10800
      TabIndex        =   2
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10440
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim menit, detik As Integer
Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000

Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
      ByVal x As Long, ByVal y As Long, _
      ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

Private Declare Function GetWindowRect Lib "user32" ( _
      ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Function ShowTitleBar(ByVal bState As Boolean)
Dim lStyle As Long
Dim tR As RECT

   GetWindowRect Me.hwnd, tR

   lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
   If (bState) Then
      Me.Caption = Me.Tag
      If Me.ControlBox Then
         lStyle = lStyle Or WS_SYSMENU
      End If
      If Me.MaxButton Then
         lStyle = lStyle Or WS_MAXIMIZEBOX
      End If
      If Me.MinButton Then
         lStyle = lStyle Or WS_MINIMIZEBOX
      End If
      If Me.Caption <> "" Then
         lStyle = lStyle Or WS_CAPTION
      End If
   Else
      Me.Tag = Me.Caption
      Me.Caption = ""
      lStyle = lStyle And Not WS_SYSMENU
      lStyle = lStyle And Not WS_MAXIMIZEBOX
      lStyle = lStyle And Not WS_MINIMIZEBOX
      lStyle = lStyle And Not WS_CAPTION
   End If
   SetWindowLong Me.hwnd, GWL_STYLE, lStyle
   SetWindowPos Me.hwnd, _
       0, tR.Left, tR.Top, _
       tR.Right - tR.Left, tR.Bottom - tR.Top, _
       SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED

   Me.Refresh

End Function

Private Sub Form_Load()
WindowsMediaPlayer1.URL = "D:\Chessmaster Grand Finale Project\02 - The First Noel.mp3"
ShowTitleBar False
ProgressBar1.Max = 110
ProgressBar1.Min = 0
ProgressBar1.Value = 0
Timer1.Interval = 130
menit = 3
detik = 9
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value < 110 Then
ProgressBar1.Value = ProgressBar1.Value + 1
End If
Label1.Caption = ProgressBar1.Value
If ProgressBar1.Value > 100 Then
Label1.Caption = "Completed!"
Label2.Visible = False
End If
If ProgressBar1.Value = 110 Then
Chess.Visible = True
Welcome.Visible = False
Timer1.Interval = 0
End If
If ProgressBar1.Value = 1 Then
Label3 = "Welcome to RagiChess-Where you will stay -Ganteng- Here!"
End If
If ProgressBar1.Value = 40 Then
Label3 = "Created by Iskandar Setiadi and VB Team 2011"
End If
If ProgressBar1.Value = 60 Then
Label3 = "Directed by Wuragil Darmoko, Sponsored by Microsoft Corp."
End If
If ProgressBar1.Value = 75 Then
Label3 = "Loading Components"
End If
If ProgressBar1.Value = 80 Then
Label3 = "Installing Graphics"
End If
If ProgressBar1.Value = 86 Then
Label3 = "Loading Interface Data"
End If
If ProgressBar1.Value = 92 Then
Label3 = "Loading Interface Data... Completed!"
End If
If ProgressBar1.Value = 95 Then
Label3 = "Enjoy the RagiChess!"
End If
End Sub

Private Sub Timer2_Timer()
If detik > 0 Then
detik = detik - 1
End If
If detik = 0 And menit = 0 Then
WindowsMediaPlayer1.URL = "D:\Chessmaster Grand Finale Project\02 - The First Noel.mp3"
menit = 3
detik = 9
End If
If detik = 0 Then
menit = menit - 1
detik = 60
End If
End Sub
