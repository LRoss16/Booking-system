VERSION 5.00
Begin VB.Form frmMembers 
   Caption         =   "Lewis' Football-Club Membership"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrintMembers 
      Caption         =   "Print Membership List"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2280
      TabIndex        =   17
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdDisplayMembers 
      Caption         =   "Display Members"
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Add Member"
      Height          =   735
      Left            =   5880
      TabIndex        =   15
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame fraDelete 
      Height          =   1935
      Left            =   3600
      TabIndex        =   12
      Top             =   2160
      Width           =   4695
      Begin VB.TextBox txtMembershipIDDelete 
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Membership Number"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ListBox lstMembers 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Frame fraAdd 
      Height          =   3615
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   4575
      Begin VB.ListBox lstCategory 
         Height          =   450
         ItemData        =   "frmMembers.frx":0000
         Left            =   1920
         List            =   "frmMembers.frx":000A
         TabIndex        =   10
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtMemberIDAdd 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Membership Category"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Membership Number"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Surname"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.OptionButton optDelete 
      Caption         =   "Delete Member"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton optAdd 
      Caption         =   "Add Member"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   1215
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisplayMembers_Click()
'If selected it will display the surname, first name, membership number and category of memberhsip i.e. Senior or Junior of all the members in a list box'

Dim FullName As String
Dim Category As String
Dim MemberID As String
Dim OneMember As MemberType

lstMembers.Clear

Open FileName For Random As #1 Len = Len(OneMember)
Do While Not EOF(1)

Get #1, , OneMember 'Program reads one record fromt the members file'
If OneMember.Deleted = "N" Then 'The record has not been deleted'
FullName = Left(UCase(OneMember.Surname) & " " & OneMember.FirstName, 24)

MemberID = OneMember.MemberID
If OneMember.Cateogry = "S" Then

Category = "Senior"

Else

Category = "Junior"

End If

lstMembers.AddItem FullName & MemberID & " " & Category

End If
Loop
Close #1
cmdPrintMembers.Enabled = True 'This can only print the report after the list box displays all of the members' details since it takes the details from the list box instead of the file'

End Sub

Private Sub cmdOK_Click()
'Processes either the adding or deletion of a member'

Dim MemberID As String
Dim MemberDeleted As Boolean 'Returns value from deleted member'
Dim Duplicate As Boolean 'Is true if the membership number is already being used'
Dim Response 'Reply when asked to confirm the deletion'
Dim OneMember As MemberType

If cmdOK.Caption = "Add Member" Then 'A new member is added;
If Len(txtMemberIDDAdd.Text) = 6 Then 'An ID has to be 6 characters long'
MemberID = txtMemberIDAdd.Text
Duplicate = CheckDuplicateMemberID(MemberID) 'Program checks to see if that membership number is already being used'

If Not Duplicate Then 'The membership number does not alreayd exist so is able to be used'

If (txtSurname.Text <> "") And (txtFirstName.Text <> "") And (lstCategory.Text <> "") Then

OneMember.MemberID = txtMemberIDAdd.Text 'The details of the member are collected into one record'

OneMember.Surname = txtSurname.Text
OneMember.FirstName = txtFirstName.Text

If lstCategory.Text = "Senior" Then
OneMember.Cateogry = "S"
Else
OneMember.Cateogry = "J"
End If

OneMember.Deleted = "N"
Call AddMember(OneMember)
txtMemberIDAdd.Text = "" 'A new record is added to the file'
txtSurname.Text = "" 'Members details are cleared'
txtFirstName.Text = ""
Else
MsgBox ("You have not filled in all details of this member")
End If
Else
MsgBox ("Membership Number & MemberID has been used. Enter a different one") 'The user has entered a membership number that is already being used'

txtMemberIDAdd.SetFocus
End If
Else
MsgBox ("You must enter a membership number with six characters")
txtMemberIDAdd.SetFocus
End If

Else

MemberID = txtMemberIDDelete.Text
If MemberID = "" Then 'A membership number has not been entered'
MsgBox ("You have not entered a membership number")
Else
Response = MsgBox("Confirm you want to delete this member?", vbYesNo)

If Response = 6 Then 'The user has confrimed to delete member'
MemberDeleted = DeleteMember(MemberID) 'This is true if the deletion is successful'

txtMemberIDDelete.Text = ""
If Not MemberDeleted Then 'The deletion was not successful'
MsgBox "Member not deleted. Membership number and MemberID does not exist", vbCritical

End If
End If
End If
End If

End Sub

Private Sub cmdPrintMembers_Click()
'This prints all of the members' details and gives the total number of juniors and seniors'
'The details are taken from the list box instead of the file'

Dim Number_of_Seniors As Integer
Dim Number_of_Juniors As Integer
Dim Category As String
Dim OneMemberDetails As String
Dim Number_of_Lines As Integer
Dim Index As Integer

Number_of_Seniors = 0
Number_of_Juniors = 0
Number_of_Lines = 0

Printer.Print 'Prints a blank page'
Printer.Print
Printer.Font.Name = "Courier" 'This is used to format the output in vertical columns'
Printer.FontSize = 16 'The size of the font'
Printer.Print "Page" & Printer.Page 'Prints the page number'
Printer.Print
Printer.Print

For Index = 0 To lstMembers.ListCount - 1 'This processes each item in the list box'
OneMemberDetails = lstMembers.List(Index) 'Current item in the list box'
Printer.Print OneMemberDetails

If InStr(OneMemberDetails, "Senior") <> 0 Then 'This returns the value zero if the search string is not present'

Number_of_Seniors = Number_of_Seniors + 1
Else
Number_of_Juniors = Number_of_Juniors + 1
End If

Number_of_Lines = Number_of_Lines + 1
If Number_of_Lines = 50 Then
Printer.NewPage 'Prints on a new page after 50 members'
Number_of_Lines = 0

Printer.Print
Printer.Print
Printer.Print "Page" & Printer.Page
Printer.Print
Printer.Print

End If

Next Index

Printer.Print
Printer.Print "Total number of Seniors " & Number_of_Seniors 'Prints how many senior members there are'
Printer.Print "Total number of juniors" & Number_of_Juniors 'Prints how many junior members there are'
Printer.Print
Printer.Print "Total amount " & Number_of_Seniors + Number_of_Juniors 'Prints the total amount of members there are'
Printer.EndDoc

End Sub

Private Sub Form_Load()

Dim FileName As String

FileName = App.Path & "\Members.dat"

End Sub

Private Sub optAdd_Click()

fraDelete.Visible = False 'Hides the frame with controls for deleting a member'
fraAdd.Visible = True 'Shows the frame with controls for adding a member'
cmdOK.Caption = "Add Member" 'Caption for command button changes to Add Member'
End Sub

Private Sub optDelete_Click()

fraAdd.Visible = False 'Hides the frame with controls for adding a member'
fraDelete.Visible = True 'Shows the frame with controls for deleting a member'
cmdOK.Caption = "Delete Member" 'Caption for command button changes to Delete Member'

End Sub

Private Function CheckDuplicateMemberID(ByVal MemberID As String) As Boolean
'Returns as true if the MemberID requested is already in use, otherwise comes back false'

Dim Found As Boolean
Dim OneMember As MemberType

Found = False
Open FileName For Random As #1 Len = Len(OneMember)
Do While (Not EOF(1)) And (Found = False) 'This will keep looping until a duplicate MemberID is found or the end of the file has been reached'

Get #1, , OneMember 'A second parameter is not necessary so 2 commas are used instead'

If MemberID = OneMember.MemberID Then
Found = True
End If
Loop
CheckDuplicateMemberID = Found 'Returns either true or false from the function'
Close #1


End Function


Private Function FindDeletedMember() As Integer
'This will find the first record in the file that has been deleted'
'A deleted record gets the deleted field set to "Y"'
'Returns the record number of the first deleted record or will return with the value 0 if there are none'

Dim Found As Boolean
Dim RecordNumber As Integer
Dim OneMember As MemberType

Found = False
RecordNumber = 0

Open FileName For Random As #1 Len = Len(OneMember)
Do While (Not EOF(1)) And (Found = False)

RecordNumber = RecordNumber + 1 'Goes to next record'
Get #1, RecordNumber, OneMember 'Reads the record from the file'
If OneMember.Deleted = "Y" Then 'Checks if record has been deleted'

Found = True
End If

Loop
If Found Then
FindDeletedMember = RecordNumber 'Returns number if there is a deleted record'
Else
FindDeletedMember = 0 'Returns value 0 if there is isn't'

End If
Close #1



End Function

Private Sub AddMember(ByRef OneMember As MemberType)
'Stores record in Members File. Calls the FindDeletedMember function to get first deleted record and uses the space left for the new record'
'If there are no deleted records it creates a new one'

Dim NumberOfRecords As Integer
Dim DeletedRecordNumber As Integer

DeletedRecordNumber = FindDeletedNumber
Open FileName For Random As #1 Len = Len(OneMember)
If DeletedRecordNumber <> 0 Then 'There is a space available from a deleted record'

Put #1, DeletedRecordNumber, OneMember 'Stores new record in space available'
Else 'No deleted records'
NumberOfRecords = LOF(1) / Len(OneMember) 'Calculates number of records'
Put #1, NumberOfRecords + 1, OneMember 'Creates a new one'
End If
Close #1
Call cmdDisplayMember_click 'This is called so the new member can appear in the list box'


End Sub

Private Function DeleteMember(ByVal MemberID As String) As Boolean
'Deletes record with Membership Number and MemberID from the members file'
'Sets deleted field to "Y"'
'Returns True if the record is successfully deleted, otherwise returns False'

Dim OneMember As MemberType
Dim RecordNumber As Integer
Dim Found As Boolean

RecordNumber = 0
Found = False
Open FileName For Random As #1 Len = Len(OneMember)
Do While (Not EOF(1)) And (Not Found)
RecordNumber = RecordNumber + 1
Get #1, RecordNumber, OneMember 'Reads one record from file'
If OneMember.MemberID = MemberID Then 'Checks if it is requires record'
If OneMember.Deleted = "Y" Then 'Member that had been previousl deleted'
MsgBox "This member does not exist"
Else
OneMember.Deleted = "Y"
Found = True
End If
End If
Loop
If Not Found Then
DeleteMember = False 'There is  no record in the file for deleting'
Else
DeleteMember = True 'A record was deleted'
End If
Put #1, RecordNumber, OneMember 'Writes record to the file'
Close #1
Call cmdDisplayMember_click 'This is called so that the deleted member is removed from the list'



End Function
