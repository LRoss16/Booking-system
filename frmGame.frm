VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Lewis' Football- Start and finish a game"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Start Game"
      Height          =   735
      Left            =   5520
      TabIndex        =   31
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Frame fraFinish 
      Height          =   4935
      Left            =   1080
      TabIndex        =   13
      Top             =   1560
      Width           =   6135
      Begin VB.TextBox txtMinutes 
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtHours 
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtCostOfGame 
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtFinishTime 
         Height          =   405
         Left            =   2520
         TabIndex        =   25
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtStartTimeFinish 
         Height          =   375
         Left            =   2520
         TabIndex        =   24
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtCategoryFinish 
         Height          =   405
         Left            =   2520
         TabIndex        =   23
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtMemberNameFinish 
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboPitchNumberFinish 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Minutes"
         Height          =   375
         Left            =   5160
         TabIndex        =   30
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Hours"
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Cost Of Game"
         Height          =   735
         Left            =   360
         TabIndex        =   20
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Playing Time"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Finish Time"
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Start Time"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Membershp Category"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Member's Name"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label label6 
         Caption         =   "Pitch Number"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCategoryStart 
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame fraStart 
      Height          =   3495
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   6015
      Begin VB.TextBox txtStartTimeStart 
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ComboBox cboPitchNumberStart 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtMemberNameStart 
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtMemberIDStart 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Start Time"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Pitch Number"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Membership Category"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Member's Name"
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Membership Number"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.OptionButton optFinish 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton optStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Value           =   -1  'True
      Width           =   975
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboPitchNumberFinish_Click()
'Call functions to retrieve the player's details and the start time on the pitch from the file'
'Displays details and then calls function NumberOfMinutes to calculate how long the game was in minutes'
'Then calculates and displays game time in hours and minutes and calls a functon to calculate the cost of the game'


Dim Pitch As Integer
Dim MinutesPlayed As Integer
Dim Hours As Integer
Dim Minutes As Integer
Dim OneGame As GameType
Dim OneMember As MemberType
Dim MembershipNumber As String
Dim Category As String

Pitch = cboPitchNumberFinish.Text
OneGame = GetRecordFromCurrentGamesFile(Pitch) 'Program retrives details of the game that has just finished form the current games file'

MembershipNumber = OneGame.MemberID
OneMember = GetMemberByMemberID(MembershipNumber) 'Rerieves the player's details from the members file'

If OneMember.Cateogry = "S" Then
txtCategoryFinish.Text = "Senior" 'Senior member'

Else

txtCategoryFinish.Text = "Junior" 'Junior member'

End If

FinishTime = Time()
txtMemberNameFinish.Text = RTrim(OneMember.FirstName) & " " & UCase(OneMember.Surname)
txtStartTimeFinish.Text = basTimeFunctions.DisplayTime(OneGame.StartTime)
txtFinishTime.Text = basTimeFunctions.DisplayTime(FinishTime)

MinutesPlayed = basTimeFunctions.NumberOfMinutes(FinishTime, OneGame.StartTime)

If MinutesPlayed >= 60 Then 'Program calculates the number of hours played'
Hours = MinutesPlayed \ 60 ' "\" calculates an integer result'
Else
Hours = 0
End If
Minutes = MinutesPlayed Mod 60 'Program calculates the number of minutes played'
txtHours.Text = Hours
txtMinutes.Text = Minutes
Category = txtCategoryFinish.Text

CostOfGame = CalculateCostOfGame(MinutesPlayed, Category)
txtCost.Text = Format(CostOfGame, "Currency") 'Program calculates the cost of the game and then displays it'

Option Explicit
'Global Variables'
Dim FinishTime As String
Dim CostOfGame As String

End Sub

Private Sub cmdOK_Click()
'Completes the processing of a new game or finished game'

Dim OneGame As GameType
Dim StartTime As Date
Dim Pitch As Integer

If cmdOK.Caption = "Start Game" Then 'A new game is being started'

If cboPitchNumberStart.Text <> "" Then 'Checks to see if a pitch has been selected'
Pitch = cboPitchNumberStart.Text

Call StoreCurrentGame(Pitch) 'Called to store the game on the file'
Call UpdatePitchDisplay(Pitch) 'Called to change PitchNumber to red'

Else
MsgBox "You have not selected a pitch number"

End If

Else 'A game is being finished'

If cboPitchNumberFinish.Text <> "" Then
Pitch = cboPitchNumberFinish.Text

OneGame = GetRecordFromCurrentGamesFile(Pitch)
StartTime = OneGame.StartTime

Call ResetGameInCurrentGamesFile(Pitch) 'Called to reset the pitch that has been used'
Call UpdatePitchDisplay(Pitch) 'Called to change PitchNumber to green'
Call StoreGameInDailyGamesFile(Pitch, StartTime) 'Called to store pitch used and start time in the file'

Else

MsgBox "You must select a pitch number"

End If

Call ListPitchesAvailable 'Called to put pitch numbers in combo box'
End Sub

Private Sub optFinish_Click()

fraFinish.Visible = True  'Shows frame with controls for finishing a game'
fraStart.Visible = False 'Hides frame with controls for starting a game'
 
cmdOK.Caption = "Finish Game" 'The caption for command box chnages to Finish Game'

Call ListPitchesAvailable 'Call subroutine of pitches available'
End Sub

Private Sub optStart_Click()

fraStart.Visible = True 'Shows frame with controls for starting a game'
fraFinish.Visible = False 'Shows frame with controls for finishing a game'
cmdOK.Caption = "Start Game" 'The caption for command box changes to Start Game'

Call ListPitchesAvailable 'Calls subroutine of pitches available'
End Sub

Private Function FindMemberByMemberID(ByVal MemberID As String) As Integer

'This function searches the members file for the membership number and MembeID'
'Returns the file record number if it exists, if it doesn't it returns the value zero'

Dim RecordNumber As Integer
Dim OneMember As MemberType
Dim FileName As String
Dim Found As Boolean

RecordNumber = 0
Found = False
FileName = App.Path & "\Members.dat"
Open FileName For Random As #1 Len = Len(OneMember)
Do While (Not EOF(1)) And (Found = False)
RecordNumber = RecordNumber + 1 'Goes to the next record in the file'
Get #1, RecordNumber, OneMember 'Reads the next record'

If OneMember.MemberID = MemberID Then 'Checks to see if it is the MemberID being searched for'

Found = True
End If
Loop
If Found Then
FindMemberByMemberID = RecordNumber 'Returns the record number if it is the MemberID being searched for'
Else
FindMemberByMemberID = 0 'Returns the value zero if it isn't the MemberID being searched for'
End If
Close #1


End Function

Private Function GetMemberByRecordNumber(ByVal RecordNumber As Integer) As MemberType
'Returns the record from the Members.dat file at the RecordNumber position'

Dim OneMember As MemberType
Dim FileName As String

FileName = App.Path & "\Members.dat"
Open FileName For Random As #1 Len = Len(OneMember)
Get #1, RecordNumber, OneMember
Close #1
GetMemberByRecordNumber = OneMember


End Function

Private Sub txtMemberIDStart_LostFocus()
'Calls FindMemberByMemberID to make sure when starting a game that the membership number entered actually exists'
'If it does exist it calls GetMemberByRecordNumber to retrieve the record and then displays the member's name and their category of membership i.e. senior or junior'

Dim MemberID As String
Dim RecordNumber As Integer
Dim OneMember As MemberType

MemberID = txtMemberIDStart.Text

If MemberID <> "" Then 'Checks to see if a number has been entered'
RecordNumber = FindMemberByMemberID(MemberID)
If RecordNumber = 0 Then 'The membership number does not exist'

MsgBox "Membership Number" & MemberID & "does not exist"
txtMemberIDStart.SetFocus
Else 'The membership number does exist'

OneMember = GetMemberRecordNumber(RecordNumber) 'Program retrieves the record from the file'
txtMemberNameStart.Text = RTrim(OneMember.FirstName) & " " & UCase(OneMember.Surname)

cboPitchNumberStart.Enabled = True 'Program will now allow the user to select a pitch number to book'

If OneMember.Cateogry = "S" Then
txtCateoryStart.Text = "Senior" 'They are a senior member'
Else
txtCategoryStart.Text = "Junior" 'They are a junior member'

End If
End If
End If

End Sub

Private Sub ListPitchesAvailable()
'The combo boxes are filled with the appropriate pitch numbers'
'If a game is being started then only green pitches are listed'
'If finishing a game then only red pitches are listed'

Dim Index As Integer

If cmdOK.Caption = "Start Game" Then 'A new game is being started'
cboPitchNumberStart.Clear

For Index = 1 To MaxPitches

If frmMain.lblPitches(Index).BackColor = vbGreen Then 'Each pitch is checked on the main form and if it is green the number is displayed'
cboPitchNumberStart.AddItem Index

End If

Next Index

Else 'A game is now finished'

cbopPitchNumberFinish.Clear

For Index = 1 To MaxPitches

If frmMain.lblPitches(Index).BackColor = vbRed Then 'Each pitch is checked on the main form and if it is red the number is displayed'
cboPitchNumberFinish.AddItem Index

End If
Next Index


End Sub

Private Sub cboPitchNumberStart_Click()
'The starting time for the new game is displayed'

Dim StartTime As String

StartTime = basTimeFunctions.ShortenTime(Time())

If Hour(Time()) > 11 Then
txtStartTimeStart.Text = StartTime & " " & "PM" 'The game has started in the afternoon or evening'
Else
txtStartTimeStart.Text = StartTime & " " & "AM" 'The game has started in the morning'

End If

End Sub

Public Sub StoreCurrentGame(ByVal PitchNumber As Integer)
'This stores the details of one game in CurrentGames.dat file'

Dim OneGame As GameType
Dim FileName As String
'It gets the details from the form's controls and stores them in a record'

OneGame.MemberID = txtMemberIDStart.Text
OneGame.PitchID = PitchNumber
OneGame.StartTime = Time()
OneGame.Occupied = "Y"

'Writes the record to the file'

FileName = App.Path & "\CurrentGames.dat"
Open FileName For Random As #1 Len = Len(OneGame)
Put #1, PitchNumber, OneGame 'Goes straight to the record that is required'
Close #1

End Sub

Private Sub UpdatePitchDisplay(ByVal PitchNumber As Integer)
'If a new game has started the PitchNumber is changed from green to red and the starting time is displayed'
'If a game is being finished the PitchNumber is changed from red to green and the starting time is removed'

If cmdOK.Caption = "Start Game" Then 'A new game is being started'
frmMain.lblPitches(PitchNumber).BackColor = vbRed 'The PitchNumber is changed to red'

If Hour(Time()) > 11 Then
frmMain.lblStartTimes(PitchNumber).Caption = basTimeFunctions.ShortenTime(Time()) & " " & "PM" 'Game has started in afternoon or evening'
Else
frmMain.lblStartTimes(PitchNumber).Caption = basTimeFunctions.ShortenTime(Time()) & " " & "AM" 'Game has started in the morning'

End If

txtMemberIDStart.Text = "" 'Text boxes are cleared for the next game'
txtMemberNameStart.Text = ""
txtCategoryStart.Text = ""
txtStartTimeStart.Text = ""

Else 'A game being finished'

frmMain.lblPitches(PitchNumber).BackColor = vbGreen 'The PitchNumber is changed to green'
frmMain.lblStartTimes(PitchNumber).Caption = " " 'The starting time is removed'

txtMemberNameFinish.Text = " " 'The text boxes are cleared'
txtCategoryFinish.Text = " "
txtStartTimeFinish.Text = " "
txtFinishTime.Text = " "
txtHours.Text = " "
txtMinutes.Text = " "
txtCost.Text = " "

End If

End Sub

Private Function GetMemberByMemberID(ByVal MemberID As String) As MemberType
'Returns the record from the members.dat file with membership number and MemberID'

Dim Found As Boolean
Dim OneMember As MemberType
Dim FileName As String

FileName = App.Path & "\Members.dat"
RecordNumber = 0

Open FileName For Random As #1 Len = Len(OneMember)
Do While Not Found 'Loops until the record is found in the file'

Get #1, , OneMember

If OneMember.MemberID = MemberID Then 'Checks if the membership number is the same as the one being searched'

Found = True

End If
Loop

GetMemberByMemberID = OneMember 'Returns the record from the function'
Close #1


End Function

Private Function CalculateCostOfGame(ByVal MinutesPlayed As Integer, ByVal Category As String) As Currency
'Calculates the cost of one game'

If Category = "Senior" Then
CaluclateCostOfGame = (SeniorRate * MinutesPlayed) / 100 'Cost for a senior member for one game'

Else
CalculateCostOfGame = (JuniorRate * MinutesPlayed) / 100 'Cost for a junior member for one game'

End If


End Function

Private Sub ResetGaneInCurrentGamesFile(ByVal PitchNumber As Integer)
'Updates the record in the current games file for PitchNumber by setting the occupied field to "N"'

Dim FileName As String
Dim OneGame As GameType

FileName = App.Path & "\CurrentGames.dat"
Open FileName For Random As #1 Len = Len(OneGame)

Get #1, PitchNumber, OneGame

OneGame.Occupied = "N"

Put #1, PitchNumber, OneGame

Close #1

End Sub


Private Sub StoreGameInDailyGamesFile(ByVal PitchNumber As Integer, ByVal StartTime As String)
'Record of the finished game is stores in the daily games file'

Dim NumberOfRecords As Integer
Dim FileName As String
Dim OneFinishedGame As GameFinishedType

FileName = App.Path & "\DailyGames.dat"

OneFinishedGame.PitchID = PitchNumber
OneFinishedGame.StartTime = StartTime
OneFinishedGame.FinishTime = FinishTime 'Uses the two global variables'
OneFinishedGame.Cost = CostOfGame

Open FileName For Random As #1 Len = Len(OneFinishedGame)

NumberOfRecords = LOF(1) / Len(OneFinishedGame)

Put #1, NumberOfRecords + 1, OneFinishedGame

Close #1

End Sub

Private Function GetRecordFromCurrentGamesFile(ByVal PitchNumber As Integer) As GameType
'Returns the record for the current game for PitchNumber'

Dim FileName As String
Dim OneGame As GameType

FileName = App.Path & "\CurrentGames.dat2"

Open FileName For Random As #1 Len = Len(OneGame)

Get #1, PitchNumber, OneGame  'Goes straight to record, retrieves it and returns it from the function'
GetRecordFromCurrentGamesFile = OneGame
Close #1

End Function
