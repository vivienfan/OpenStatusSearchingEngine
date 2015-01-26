VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "GR Location"
   ClientHeight    =   4485
   ClientLeft      =   7020
   ClientTop       =   3285
   ClientWidth     =   6855
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   6855
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6615
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3480
      Width           =   6615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = Form1.gr_path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
Text1.Text = Dir1.path
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
Text1.Text = File1.path
End Sub

Private Sub file1_click()
Text1.Text = File1.path + "\" + File1.filename
End Sub

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub OK_Click()
Form1.gr_path = Text1.Text
Unload Me
End Sub

