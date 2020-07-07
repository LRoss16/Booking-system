Attribute VB_Name = "basDeclarations"
Public Const MaxPitches = 20 'There are 20 football pitches available'

Public Type GameType 'Data type for record to store one current game'
MemberID As String * 6 '6 bytes'
PitchID As Integer '2 bytes'
StartTime As Date '8 bytes'
Occupied As String * 1 '1 byte'
End Type 'One record is 17 bytes'


Public Type GameFinishedType 'Data type for a record for a game that is finished'

PitchID As Integer '2 bytes'
StartTime As Date '8 bytes'
FinishTime As Date '8 bytes'
Cost As Currency '8 bytes'
End Type 'One records is 26 bytes'


Public Type MemberType 'Data type for a record to store details for one member'

MemberID As String * 6 'bytes'
Surname As String * 14 '14 bytes'
FirstName As String * 20 '20 bytes'
Cateogry As String * 1 '1 byte'
Deleted As String * 1 '1 bytes'
End Type 'One record is 42 bytes'

Public SeniorRate As Single
Public JuniorRate As Single

Public Sub Main()
Dim FileName As String

FileName = App.Path & "\Costs.txt"

Open FileName For Input Access Read As #1
Input #1, SeniorRate, JuniorRate
Close #1

frmMain.Show

End Sub
