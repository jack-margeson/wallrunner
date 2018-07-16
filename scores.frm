VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmScores 
   BackColor       =   &H80000015&
   Caption         =   "Scores"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H80000015&
      Caption         =   "Exit"
      Height          =   375
      Left            =   2220
      TabIndex        =   4
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H80000015&
      Caption         =   "Return to Main Menu"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   5280
      Width           =   1815
   End
   Begin VB.ListBox lstPoints 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   4155
      Left            =   2160
      TabIndex        =   2
      Top             =   1020
      Width           =   1935
   End
   Begin VB.ListBox lstName 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cmdOpen 
      Left            =   3720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   3180
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "High Scores"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim names(1 To 10) As String
Dim points(1 To 10) As String
Dim verify As Boolean
Dim newname As String
Dim replace As Integer
Dim failed As Integer
Public SubmitScore As Integer
Public score As Integer

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdReturn_Click()
Unload frmScores
frmMenu.Show
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim line As String

'Loads points into array.
Open App.Path & "\highscores\points.txt" For Input As #1
    For i = 1 To 10
        If EOF(1) Then
            Exit For
        End If
        Line Input #1, line
        line = Trim(line)
        If Len(line) <> 0 Then
            points(i) = line
            lstPoints.AddItem line
        End If
    Next i
Close #1

'Loads names into array.
Open App.Path & "\highscores\names.txt" For Input As #1
    For i = 1 To 10
        If EOF(1) Then
            Exit For
        End If
        Line Input #1, line
        line = Trim(line)
        If Len(line) <> 0 Then
            points(i) = line
            lstName.AddItem line
        End If
    Next i
Close #1

'The first part checks if you're entering the high score menu after you play a game (SubmitScore) or from the menu.
'The second part checks if the score trying to be submitted is valid and is larger than a previous score. (failed)
'The third part sets the verify flag to being true so that the lowest score can be replaced.
If SubmitScore = 1 Then
    For i = 0 To 10
        If SubmitScore = 1 Then
            If Val(score) <= Val(lstPoints.List(i)) Then
                failed = failed + 1
                If failed = 10 Then
                    SubmitScore = 0
                    MsgBox ("Your score does not qualify for a record on the leaderboards... sorry!"), vbOKOnly, "Wall Runner"
                    verify = False
                    Unload frmScores
                    frmScores.Show
                End If
            ElseIf Val(score) > Val(lstPoints.List(i)) Then
                SubmitScore = 0
                verify = True
                Dim answer As String
                answer = InputBox("Please enter your initials!", "Wall Runner", "JIM")
                newname = answer
                replace = Val(i)
            End If
        End If
    Next i
End If

If verify = True Then
    'Moves everything down a peg. 0-9. This is gross looking - and theres probably a better way.
    If replace = 0 Then
        lstPoints.List(replace + 9) = lstPoints.List(replace + 8)
        lstPoints.List(replace + 8) = lstPoints.List(replace + 7)
        lstPoints.List(replace + 7) = lstPoints.List(replace + 6)
        lstPoints.List(replace + 6) = lstPoints.List(replace + 5)
        lstPoints.List(replace + 5) = lstPoints.List(replace + 4)
        lstPoints.List(replace + 4) = lstPoints.List(replace + 3)
        lstPoints.List(replace + 3) = lstPoints.List(replace + 2)
        lstPoints.List(replace + 2) = lstPoints.List(replace + 1)
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 9) = lstName.List(replace + 8)
        lstName.List(replace + 8) = lstName.List(replace + 7)
        lstName.List(replace + 7) = lstName.List(replace + 6)
        lstName.List(replace + 6) = lstName.List(replace + 5)
        lstName.List(replace + 5) = lstName.List(replace + 4)
        lstName.List(replace + 4) = lstName.List(replace + 3)
        lstName.List(replace + 3) = lstName.List(replace + 2)
        lstName.List(replace + 2) = lstName.List(replace + 1)
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 1 Then
        lstPoints.List(replace + 8) = lstPoints.List(replace + 7)
        lstPoints.List(replace + 7) = lstPoints.List(replace + 6)
        lstPoints.List(replace + 6) = lstPoints.List(replace + 5)
        lstPoints.List(replace + 5) = lstPoints.List(replace + 4)
        lstPoints.List(replace + 4) = lstPoints.List(replace + 3)
        lstPoints.List(replace + 3) = lstPoints.List(replace + 2)
        lstPoints.List(replace + 2) = lstPoints.List(replace + 1)
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 8) = lstName.List(replace + 7)
        lstName.List(replace + 7) = lstName.List(replace + 6)
        lstName.List(replace + 6) = lstName.List(replace + 5)
        lstName.List(replace + 5) = lstName.List(replace + 4)
        lstName.List(replace + 4) = lstName.List(replace + 3)
        lstName.List(replace + 3) = lstName.List(replace + 2)
        lstName.List(replace + 2) = lstName.List(replace + 1)
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 2 Then
        lstPoints.List(replace + 7) = lstPoints.List(replace + 6)
        lstPoints.List(replace + 6) = lstPoints.List(replace + 5)
        lstPoints.List(replace + 5) = lstPoints.List(replace + 4)
        lstPoints.List(replace + 4) = lstPoints.List(replace + 3)
        lstPoints.List(replace + 3) = lstPoints.List(replace + 2)
        lstPoints.List(replace + 2) = lstPoints.List(replace + 1)
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 7) = lstName.List(replace + 6)
        lstName.List(replace + 6) = lstName.List(replace + 5)
        lstName.List(replace + 5) = lstName.List(replace + 4)
        lstName.List(replace + 4) = lstName.List(replace + 3)
        lstName.List(replace + 3) = lstName.List(replace + 2)
        lstName.List(replace + 2) = lstName.List(replace + 1)
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 3 Then
        lstPoints.List(replace + 6) = lstPoints.List(replace + 5)
        lstPoints.List(replace + 5) = lstPoints.List(replace + 4)
        lstPoints.List(replace + 4) = lstPoints.List(replace + 3)
        lstPoints.List(replace + 3) = lstPoints.List(replace + 2)
        lstPoints.List(replace + 2) = lstPoints.List(replace + 1)
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 6) = lstName.List(replace + 5)
        lstName.List(replace + 5) = lstName.List(replace + 4)
        lstName.List(replace + 4) = lstName.List(replace + 3)
        lstName.List(replace + 3) = lstName.List(replace + 2)
        lstName.List(replace + 2) = lstName.List(replace + 1)
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 4 Then
        lstPoints.List(replace + 5) = lstPoints.List(replace + 4)
        lstPoints.List(replace + 4) = lstPoints.List(replace + 3)
        lstPoints.List(replace + 3) = lstPoints.List(replace + 2)
        lstPoints.List(replace + 2) = lstPoints.List(replace + 1)
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 5) = lstName.List(replace + 4)
        lstName.List(replace + 4) = lstName.List(replace + 3)
        lstName.List(replace + 3) = lstName.List(replace + 2)
        lstName.List(replace + 2) = lstName.List(replace + 1)
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 5 Then
        lstPoints.List(replace + 4) = lstPoints.List(replace + 3)
        lstPoints.List(replace + 3) = lstPoints.List(replace + 2)
        lstPoints.List(replace + 2) = lstPoints.List(replace + 1)
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 4) = lstName.List(replace + 3)
        lstName.List(replace + 3) = lstName.List(replace + 2)
        lstName.List(replace + 2) = lstName.List(replace + 1)
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 6 Then
        lstPoints.List(replace + 3) = lstPoints.List(replace + 2)
        lstPoints.List(replace + 2) = lstPoints.List(replace + 1)
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 3) = lstName.List(replace + 2)
        lstName.List(replace + 2) = lstName.List(replace + 1)
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 7 Then
        lstPoints.List(replace + 2) = lstPoints.List(replace + 1)
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 2) = lstName.List(replace + 1)
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 8 Then
        lstPoints.List(replace + 1) = lstPoints.List(replace)
        
        lstName.List(replace + 1) = lstName.List(replace)
        
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    ElseIf replace = 9 Then
        lstPoints.List(replace) = score
        lstName.List(replace) = newname
    End If
End If

'Saves the final outputs to their respective file for future loading.
Open App.Path & "\highscores\points.txt" For Output As #1
    For i = 0 To 10
        Print #1, lstPoints.List(i)
    Next i
Close #1

Open App.Path & "\highscores\names.txt" For Output As #1
    For i = 0 To 10
        Print #1, lstName.List(i)
    Next i
Close #1
End Sub
