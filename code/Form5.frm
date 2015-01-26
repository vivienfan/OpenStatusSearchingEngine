VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Save Excel As..."
   ClientHeight    =   4485
   ClientLeft      =   7080
   ClientTop       =   3390
   ClientWidth     =   6870
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   4485
   ScaleWidth      =   6870
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3480
      Width           =   6615
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   6375
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = "C:\Users\Vivien Fan\My Documents\^~ ^ Vi Vi Ann\Visual Basic"
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
Text1.Text = Dir1.path
End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.path
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim filename, d As String
filename = "\OpenStatus" & Date & ".xls"
MsgBox filename
Form1.newbook.SaveAs Text1.Text & filename
Form1.newbook.Close
Form1.newApp.Quit
Set Form1.newApp = Nothing
Form1.NewExcel = False
Unload Me
End Sub

