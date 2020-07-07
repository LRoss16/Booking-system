VERSION 5.00
Begin VB.Form frmUtilities 
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCosts 
      Height          =   1095
      Left            =   600
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtJuniorCost 
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSeniorCost 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Junior Rate"
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Senior Rate"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   855
      Left            =   2640
      TabIndex        =   5
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6855
      Begin VB.OptionButton optUtilities 
         Caption         =   "Back up Current Games and Daily Games Files"
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optUtilities 
         Caption         =   "Back up Members file"
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optUtilities 
         Caption         =   "Create a new Current Games File"
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   2175
      End
      Begin VB.OptionButton optUtilities 
         Caption         =   "Change cost of Games"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   1
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label lblHelp 
         Caption         =   "Enter as pounds per hour eg 70, 60.5 etc"
         Height          =   735
         Left            =   5040
         TabIndex        =   11
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Dim Index As Integer
Dim OptionChoice As Integer 'Option button selected from 1 to 4
For Index = 1 To 4 'Finds out which option has been selected'
If optUtilities(Index).Value = True Then 'This is true if the option has been selected'
OptionChoice = Index
End If
Next Index
Select Case OptionChoice
 Case 1
 Call BackupGamesFiles
  Case 2
  Call BackupMembersFile
Case 3
 Call CreateCurrentGamesFile
Case 4
Call ChangeCostOfGame
    
End Select

End Sub

Private Sub optUtilities_Click(Index As Integer)
If Index = 4 Then
fraCosts.Visible = True
lblHelp.Visible = True
End If

End Sub
Public Sub BackupGamesFiles() 'This uses the file copy statement to copy the two files into the floppy disk to back up the files'

Dim Source1 As String
Dim Source2 As String
Dim Destination1 As String
Dim Destination2 As String

Source1 = App.Path & "\CurrentGames.dat"

Destination1 = "E:\CurrentGames.dat"

FileCopy Source1, Destination1

Source2 = App.Path & "\DailyGames.dat"

Destination2 = "E:\DailyGames.dat"

FileCopy Source2, Destination2


End Sub

Public Sub BackupMembersFile() 'This uses the file copy statement to copy the file into a floppydisk to backup the file'

Dim Source As String
Dim Destination As String

Source = App.Path & "\Members.dat"

Destination = "E:\Members.dat"

FileCopy Source, Destination

End Sub


Public Sub CreateCurrentGamesFile() 'This creates a current games file with one record for each pitch'

Dim OneGame As GameType
Dim PitchNumber As Integer

Open App.Path & "\CurrentGames.dat" For Random As #1 Len = Len(OneGame)

For PitchNumber = 1 To MaxPitches 'Loops 20 times, once for each pitch'

OneGame.MemberID = "" 'This is set to blank'
OneGame.PitchID = PitchNumber 'assigns a number from 1 to 20 depending on pitch being used'
OneGame.Occupied = "N" 'This is set for the pitch to come up as not occupied'
Put #1, PitchNumber, OneGame 'This writes the record to the file'
Next PitchNumber
Close #1

End Sub

Public Sub ChangeCostOfGame() 'This stores the new costs for seniors and juniors in the costs file'

Dim FileName As String
Dim SeniorCost As String
Dim JuniorCost As String

SeniorCost = txtSeniorCost.Text 'This gets the new price for seniors'
JuniorCost = txtJuniorCost.Text 'This gets the new price for juniors'

If (Not IsNumeric(SeniorCost)) Or (Not IsNumeric(JuniorCost)) Then
MsgBox ("One or both the rate are not numbers. Please re-enter")

Else

FileName = App.Path & "\Costs.txt"
Open FileName For Output As #1
Write #1, SeniorCost, JuniorCost 'Writes the new prices to the file'
Close #1
End If


End Sub
