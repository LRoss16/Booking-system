VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGame 
      Caption         =   "Game"
      Height          =   735
      Left            =   9720
      TabIndex        =   46
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdMembership 
      Caption         =   "Membership"
      Height          =   735
      Left            =   9720
      TabIndex        =   45
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdUtilities 
      Caption         =   "Utilities"
      Height          =   735
      Left            =   9720
      TabIndex        =   44
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   855
      Left            =   7920
      TabIndex        =   43
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Pitch Usage Report"
      Height          =   855
      Left            =   840
      TabIndex        =   42
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   5775
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   9255
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   39
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   240
         TabIndex        =   38
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   2280
         TabIndex        =   36
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   2280
         TabIndex        =   35
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   2280
         TabIndex        =   34
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   4320
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   4320
         TabIndex        =   32
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   6720
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   4320
         TabIndex        =   30
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   2280
         TabIndex        =   29
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   4320
         TabIndex        =   28
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   4320
         TabIndex        =   27
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   2280
         TabIndex        =   26
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   6720
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   18
         Left            =   6720
         TabIndex        =   24
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   6720
         TabIndex        =   23
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblPitches 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   19
         Left            =   6720
         TabIndex        =   22
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   19
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1080
         TabIndex        =   18
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1080
         TabIndex        =   17
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3240
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3240
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3240
         TabIndex        =   14
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3240
         TabIndex        =   13
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3240
         TabIndex        =   12
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   5160
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   5160
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   5160
         TabIndex        =   9
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   5160
         TabIndex        =   8
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   5160
         TabIndex        =   7
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   7680
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   7680
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   7680
         TabIndex        =   4
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   7680
         TabIndex        =   3
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblStartTimes 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   7680
         TabIndex        =   2
         Top             =   4200
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "                              Lewis' Football Pitch Hire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGame_Click()
frmGame.Show

End Sub

Private Sub cmdMembership_Click()
frmMembers.Show

End Sub

Private Sub cmdPrintReport_Click()

Dim Index As Integer
Dim OneFinishedGame As GameFinishedType
Dim NumberOfRecords As Integer
Dim RecordNumber As Integer
Dim Pitch As Integer
Dim GameForOnePitch As Integer 'Number of games played on a pitch'
Dim TimeForOnePitch            'Total time in minutes a pitch has been used'
Dim IncomeForOnePitch As Currency 'The total income made from a pitch'
Dim TotalIncome As Currency 'The total income from all pitches'
Dim MinutesPlayed As Integer 'Number of minutes a game has lasted'
Dim Hours As Integer
Dim Minutes As Integer
Dim FileName As String

TotalIncome = 0



Printer.Print
Printer.Print
Printer.FontSize = 16
Printer.FontBold = True
Printer.Print "Lewis' Football Pitch Hire. Pitch use for" & Date 'Report header with the current date, fontsize is 16 and is in bold'
Printer.Print
Printer.Print
Printer.FontSize = 12
Printer.FontBold = False
Printer.Print "Pitch Number."; Tab(15); "Number of Game"; Tab(30); "Total Time"; Tab(50); "Income" 'Column headinggs, the fontsize is 12 and they are not in bold'
Printer.Print

FileName = App.Path & "\DailyGames.dat"

For Pitch = 1 To MaxPitches 'Looped for each pitch'

Open FileName For Random As #1 Len = Len(OneFinishedGame)
NumberOfRecords = FileLen(FileName) / Len(OneFinishedGame)

GamesForOnePitch = 0
TimeForOnePitch = 0
IncomeForOnePitch = 0

For Index = 1 To NumberOfRecords 'Program searches the whole file for records of the current pitch'

Get #1, , OneFinishedGame 'Reads one record'

If OneFinishedGame.PitchID = Pitch Then 'The current pitch in use is found'

GamesForOnePitch = GamesForOnePitch + 1 'Adds one more game to it'

MinutesPlayed = basTimeFunctions.NumberOfMinutes(OneFinishedGame.FinishTime, OneFinishedGame.StartTime)

TimeForOnePitch = TimeForOnePitch + MinutesPlayed

IncomeForOnePitch = IncomeForOnePitch + OneFinishedGame.Cost 'Adds the income from the pitch to the total income'

End If

Next Index

TotalIncome = TotalIncome + IncomeForOnePitch

If TimeForOnePitch >= 60 Then 'The program calculates the hours played on the current pitch'
Hours = TimeForOnePitch \ 60
Else
Hours = 0
End If

Minutes = TimeForOnePitch Mod 60 'Program calculates the minutes played'
Printer.Print Tab(5); Pitch; Tab(15); GamesForOnePitch; Tab(30); Format(IncomeForOneTable, "Currency")

Close #1
Next Pitch

Printer.Print
Printer.Print Tab(30); "Total Income"; Tab(50); Format(TotalIncome, "Currency")
Printer.EndDoc

Call DeleteDailyGamesFile
End Sub

Private Sub cmdUtilities_Click()
frmUtilities.Show 'The form utilities is shown'

End Sub



Private Sub DeleteDailyGamesFile() 'Deletes the day's game files, it does not go to recycling bin'

Dim Response As Integer
Dim FileName As String

FileName = App.Path & "\DailyGames.dat"

Response = MsgBox("Delete today's games from the file?", vbYesNo)
If Response = 6 Then 'User has responded yes'
Kill FileName 'File is deleted'
End If

End Sub


Private Sub Form_Load()


Dim Index As Integer
Dim Pitch As Integer
Dim FileName As String
Dim OneGame As GameType

FileName = App.Path & "\CurrentGames.dat"
Open FileName For Random As #1 Len = Len(OneGame)
For Pitch = 1 To MaxPitches

Get #1, , OneGame

If OneGame.Occupied = "Y" Then

lblPitches(Pitch).BackColor = vbRed
lblStartTimes(Pitch).Caption = basTimeFunctions.ShortenTime(OneGame.StartTime)

End If
Next Pitch
Close #1
Call frmGame.ListPitchesAvailable





End Sub
