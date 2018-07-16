VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H80000015&
   Caption         =   "Wall Runner"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1500
      Picture         =   "menu.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   1560
      Width           =   2595
   End
   Begin VB.CommandButton cmdInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "How To Play"
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
      TabIndex        =   6
      Top             =   5820
      Width           =   2235
   End
   Begin VB.Timer tmrRun 
      Interval        =   350
      Left            =   3660
      Top             =   10920
   End
   Begin VB.PictureBox picGuyWave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   1920
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   5
      Top             =   7380
      Width           =   1920
   End
   Begin VB.CommandButton cmdPlay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "Play"
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
      Top             =   4380
      Width           =   2235
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "Exit"
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
      Top             =   6540
      Width           =   2235
   End
   Begin VB.CommandButton cmdScores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "High Scores"
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
      TabIndex        =   2
      Top             =   5100
      Width           =   2235
   End
   Begin VB.PictureBox picLeftWall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000011&
      Height          =   11415
      Left            =   0
      Picture         =   "menu.frx":38F5
      ScaleHeight     =   11415
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   0
      Width           =   1500
   End
   Begin VB.PictureBox picRightWall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000011&
      Height          =   11415
      Left            =   4140
      Picture         =   "menu.frx":7C57
      ScaleHeight     =   11415
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currentimgguy As Integer
Dim imgguy As String
Public SubmitScore As Integer

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdInfo_Click()
MsgBox ("Welcome to Wall Runner!" & vbLf & "----------------------------------" & vbLf & vbLf & "Control the runner using the left and right arrow keys." & vbLf & "Be careful though, don't hit the walls on the side or on the path!" & vbLf & "Try going for a high score -" & vbLf & "see how many walls you can clear!" & vbLf & vbLf & "Created by Jack Margeson." & vbLf & "Computer Programming Credit Flex" & vbLf & "Summer of 2018"), vbOKOnly, "Wall Runner"
End Sub

Private Sub cmdPlay_Click()
Me.Hide
SubmitScore = 0
frmScores.SubmitScore = SubmitScore
Unload frmGame
frmGame.Show
End Sub

Private Sub cmdScores_Click()
Me.Hide
SubmitScore = 0
frmScores.SubmitScore = SubmitScore
frmScores.Show
End Sub

'Timer controlled animation for the playable character.
Private Sub tmrRun_Timer()
If currentimgguy = 0 Then
    imgguy = ""
    imgguy = App.Path & "\textures\guywave2.jpg"
    picGuyWave.Picture = LoadPicture(imgguy)
    currentimgguy = 1
ElseIf currentimgguy = 1 Then
    imgguy = ""
    imgguy = App.Path & "\textures\guywave1.jpg"
    picGuyWave.Picture = LoadPicture(imgguy)
    currentimgguy = 0
End If
End Sub
