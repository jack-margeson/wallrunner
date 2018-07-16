VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmGame 
   BackColor       =   &H80000015&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wall Runner"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   11415
   ScaleWidth      =   5760
   Begin VB.PictureBox picWall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   1500
      Picture         =   "game.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   11
      Top             =   -450
      Width           =   1500
   End
   Begin VB.PictureBox picLeftWall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000011&
      Height          =   11415
      Left            =   0
      Picture         =   "game.frx":26E1
      ScaleHeight     =   11415
      ScaleWidth      =   1500
      TabIndex        =   10
      Top             =   0
      Width           =   1500
   End
   Begin VB.PictureBox picGo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   1920
      Picture         =   "game.frx":6A43
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   9
      Top             =   4080
      Width           =   1920
   End
   Begin VB.PictureBox pic3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   1920
      Picture         =   "game.frx":95A2
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   8
      Top             =   4080
      Width           =   1920
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   1920
      Picture         =   "game.frx":C0A6
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   7
      Top             =   4080
      Width           =   1920
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   1980
      Picture         =   "game.frx":180E8
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   6
      Top             =   4080
      Width           =   1920
   End
   Begin VB.CommandButton cmdSubmitScore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "Submit Score"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1740
      MaskColor       =   &H80000012&
      TabIndex        =   5
      Top             =   6540
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "Return to Menu"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1740
      MaskColor       =   &H80000012&
      TabIndex        =   4
      Top             =   7260
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CommandButton cmdPlayAgain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "Play Again"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1740
      MaskColor       =   &H80000012&
      TabIndex        =   3
      Top             =   5820
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picGuy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   1680
      MousePointer    =   2  'Cross
      Picture         =   "game.frx":1AB70
      ScaleHeight     =   2194.286
      ScaleMode       =   0  'User
      ScaleWidth      =   960
      TabIndex        =   1
      Top             =   9420
      Width           =   960
   End
   Begin VB.PictureBox picRightWall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000011&
      Height          =   11415
      Left            =   4260
      Picture         =   "game.frx":1D229
      ScaleHeight     =   11415
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      Begin VB.Timer tmrGameSpeed 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   120
      End
      Begin VB.Timer tmrGame 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   540
      End
      Begin VB.Timer tmrCountdown 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   960
         Top             =   10440
      End
      Begin VB.Timer tmrMove 
         Interval        =   1
         Left            =   540
         Top             =   10860
      End
      Begin VB.Timer tmrRun 
         Interval        =   350
         Left            =   960
         Top             =   10860
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpMusic 
      Height          =   435
      Left            =   3600
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
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
      _cy             =   767
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpGo 
      Height          =   435
      Left            =   3120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
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
      _cy             =   767
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   435
      Left            =   2640
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
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
      _cy             =   767
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   435
      Left            =   2160
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
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
      _cy             =   767
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp3 
      Height          =   435
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
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
      _cy             =   767
   End
   Begin VB.Label lblPointCounter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1215
      Left            =   2160
      TabIndex        =   12
      Top             =   1380
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "Game Over!"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   1560
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim imgguy As String
Dim currentimgguy As Integer
Dim currentpos As Integer
Dim gamestate As Boolean
Dim counterstep As Integer
Dim speed As Integer
Dim currentwallpos As Integer
Dim firstloop As Boolean
Dim GameOver As Boolean

Public score As Integer
Public SubmitScore As Integer

'Sets positons of playable character. Makes it easier to work with.
Dim pos1 As Integer
Dim pos2 As Integer
Dim pos3 As Integer
Dim pos4 As Integer

Private Sub cmdExit_Click()
Me.Hide
frmMenu.Show
End Sub

'Play the game again after loss.
Private Sub cmdPlayAgain_Click()
Unload frmGame
frmGame.Show
End Sub

'Submit scores.
Private Sub cmdSubmitScore_Click()
If GameOver = True Then
    SubmitScore = 1
    frmScores.SubmitScore = SubmitScore
    frmScores.score = score
    Unload frmGame
    frmScores.Show
End If
End Sub

'Sets up the game.
Private Sub Form_Activate()
gamestate = False
GameOver = True
score = 0
tmrGame.Enabled = False
tmrGameSpeed.Enabled = False

lblGameOver.Visible = False
cmdPlayAgain.Visible = False
cmdExit.Visible = False
cmdPlayAgain.Enabled = False
cmdExit.Enabled = False
cmdSubmitScore.Visible = False
cmdSubmitScore.Enabled = False

pic1.Visible = False
pic2.Visible = False
pic3.Visible = False
picGo.Visible = False
counterstep = 1
    pic3.Visible = True
    wmpMusic.Enabled = True
    wmp3.settings.volume = 10
    wmp2.settings.volume = 10
    wmp1.settings.volume = 10
    wmpGo.settings.volume = 10
    wmpMusic.settings.volume = 10
    wmpMusic.settings.setMode "loop", True
    
    wmp3.URL = App.Path & "\sounds\3.wav"
    counterstep = counterstep + 1
tmrCountdown.Enabled = True

'Sets positons of playable character. Makes it easier to work with.
pos1 = 240
pos2 = 1680
pos3 = 3120
pos4 = 4560

firstloop = True
picLeftWall.ZOrder 1
picGuy.ZOrder 1

currentpos = 2
picGuy.Top = 9600
picGuy.Left = 1680
End Sub

'Controls the countdown and facilitates the start of the game.
Private Sub tmrCountdown_Timer()

If counterstep = 1 Then
    pic3.Visible = True
    wmp3.URL = App.Path & "\sounds\3.wav"
    counterstep = counterstep + 1
ElseIf counterstep = 2 Then
    pic3.Visible = False
    pic2.Visible = True
    wmp2.URL = App.Path & "\sounds\2.wav"
    counterstep = counterstep + 1
ElseIf counterstep = 3 Then
    pic2.Visible = False
    pic1.Visible = True
    wmp1.URL = App.Path & "\sounds\1.wav"
    counterstep = counterstep + 1
ElseIf counterstep = 4 Then
    pic1.Visible = False
    picGo.Visible = True
    wmpGo.URL = App.Path & "\sounds\go.wav"
    counterstep = counterstep + 1
    'Start game operations
ElseIf counterstep = 5 Then
    wmpMusic.URL = App.Path & "\sounds\music.wav"
    picGo.Visible = False
    lblPointCounter.Visible = True
    tmrCountdown.Enabled = False
    tmrGameSpeed.Enabled = True
    tmrGame.Enabled = True
    gamestate = True
End If
End Sub

'Controls the speed of the gameplay.
Private Sub tmrGameSpeed_Timer()
If score = 0 Then
    speed = 30
Else
    If speed <= 225 Then
        speed = (score + 1) * 15
    ElseIf speed > 225 Then
        speed = 225
        tmrGameSpeed.Enabled = False
    End If
End If
End Sub

'Controls the actual gameplay.
Private Sub tmrGame_Timer()

If firstloop = True Then
    currentwallpos = 2
    firstloop = False
End If

If picWall.Top < 11500 Then
    picWall.Top = picWall.Top + speed
    
'This is the collison testing for the 2nd and 3rd moveable lanes in the game - the actual moving walls.
    Do While 9200 < picWall.Top And picWall.Top < 10200
        If currentwallpos = 2 And currentpos = 2 Or currentwallpos = 3 And currentpos = 3 Then
            tmrGame.Enabled = False
            gamestate = False
            GameOverBySideWall
            Exit Do
        Else
            Exit Do
        End If
    Loop
End If

If picWall.Top > 11500 Then
    score = score + 1
    lblPointCounter.Caption = score
    picWall.Visible = False
                
    'Chooses what side the wall will be located next.
    currentwallpos = Int((3 - 2 + 1) * Rnd) + 2
                
    If currentwallpos = 2 Then
        picWall.Top = -450
        picWall.Left = 1500
        picWall.Visible = True
    ElseIf currentwallpos = 3 Then
        picWall.Top = -450
        picWall.Left = 2760
        picWall.Visible = True
    End If
End If
End Sub

'Control system for the playable character.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If gamestate = True Then
    'Alongside the controls, collision testing is done for the 1st and 4th moveable lanes in the game - the static walls.
    If KeyCode = 37 Then 'Left
        If currentpos = 2 Then
            picGuy.Left = pos1
            currentpos = 1
            GameOverBySideWall
        ElseIf currentpos = 3 Then
            picGuy.Left = pos2
            currentpos = 2
        End If
    ElseIf KeyCode = 39 Then 'Right
        If currentpos = 2 Then
            picGuy.Left = pos3
            currentpos = 3
        ElseIf currentpos = 3 Then
            picGuy.Left = pos4
            currentpos = 4
            GameOverBySideWall
        End If
    End If
End If
End Sub

'Timer controlled animation for the playable character.
Private Sub tmrRun_Timer()
If gamestate = True Then
    If currentimgguy = 0 Then
        imgguy = ""
        imgguy = App.Path & "\textures\guy2.jpg"
        picGuy.Picture = LoadPicture(imgguy)
        currentimgguy = 1
    ElseIf currentimgguy = 1 Then
        imgguy = ""
        imgguy = App.Path & "\textures\guy1.jpg"
        picGuy.Picture = LoadPicture(imgguy)
        currentimgguy = 0
    End If
End If
End Sub

Sub GameOverBySideWall()
picGuy.ZOrder 0
imgguy = App.Path & "\textures\explosion.jpg"
picGuy.Picture = LoadPicture(imgguy)

gamestate = False
GameOver = True
wmpMusic.URL = ""
frmScores.score = score
tmrGame.Enabled = False
tmrGameSpeed.Enabled = False

lblGameOver.Visible = True
cmdPlayAgain.Visible = True
cmdExit.Visible = True
cmdPlayAgain.Enabled = True
cmdExit.Enabled = True
cmdSubmitScore.Visible = True
cmdSubmitScore.Enabled = True
tmrRun.Enabled = False
End Sub


