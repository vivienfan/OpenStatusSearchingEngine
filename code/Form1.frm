VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Open Cases Searching Engine"
   ClientHeight    =   4095
   ClientLeft      =   7485
   ClientTop       =   4080
   ClientWidth     =   5640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5640
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search"
      Height          =   3135
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   2175
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Open"
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Run"
         Default         =   -1  'True
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Open IR"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Open GR"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Open IR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   2895
      Begin VB.Label Label6 
         Caption         =   "0"
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Case Amount: "
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Open GR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.Label Label5 
         Caption         =   "0"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Case Amount:"
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Menu File 
      Caption         =   "File(&F)"
      Begin VB.Menu Locations 
         Caption         =   "Locations"
         Begin VB.Menu GRLocation 
            Caption         =   "GR File Location"
         End
         Begin VB.Menu IRLocation 
            Caption         =   "IR File Location"
         End
         Begin VB.Menu emp_info 
            Caption         =   "Employee Info. Location"
         End
      End
      Begin VB.Menu Divider1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help(&H)"
      Begin VB.Menu ReadMe 
         Caption         =   "Read Me"
         Enabled         =   0   'False
      End
      Begin VB.Menu Divider2 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'gr excel files
Dim grApp As Excel.Application
Dim grbook As Excel.Workbook
Dim grsheet As Excel.Worksheet
'ir excel files
Dim irApp As Excel.Application
Dim irbook As Excel.Workbook
Dim irsheet As Excel.Worksheet
'sap excel files
Dim eApp As Excel.Application
Dim ebook As Excel.Workbook
Dim esheet As Excel.Worksheet
'create new excel
Public newApp As Excel.Application
Public newbook As Excel.Workbook
Public newsheet1 As Excel.Worksheet
Public newsheet2 As Excel.Worksheet
'file path
Public gr_path As String
Public ir_path As String
Public emp_path As String
'condition variables
Public NewExcel As Boolean

Function search()
'this function returns users' command for further usage
'None -> String
Dim opt As String
If Check1.Value = 1 Then
opt = opt + "g"
End If
If Check2.Value = 1 Then
opt = opt + "i"
End If
search = opt
End Function

Function find_col(info As String, sheet As Excel.Worksheet)
'this function returns the exact column index
'(String, Excel.worksheet) -> Integer
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Finding " & info
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Dim cell As String
Dim i As Integer
i = 1
cell = LCase(Trim(sheet.Cells(1, i).Value))
While info <> cell
i = i + 1
cell = LCase(Trim(sheet.Cells(1, i).Value))
Wend
If i <> 0 Then
find_col = i
Else
MsgBox info & " is not found in the database.Empty column will be filled in instead"
find_col = 50
End If
End Function

Function status(opt As String, sheet As Excel.Worksheet)
'this function takes in a excel sheet,
'finds gr, ir qty coresponsding columns,
'and then returns an arrary of open case row numbers
'(excel.worksheet) -> (arrary)
'variables declaration
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Calling status function"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Finding open cases"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Dim gr_col, ir_col As Integer
Dim arr() As String
Dim i As Integer
Dim gr_qty, ir_qty As String
ReDim arr(0)
'initializing variables
gr_col = find_col("gr qty", sheet)
ir_col = find_col("ir qty", sheet)
arr(0) = "1"
i = 2
gr_qty = sheet.Cells(2, gr_col)
ir_qty = sheet.Cells(2, ir_col)
'while loop
Select Case opt
Case "gr"
While gr_qty <> ""
If Val(gr_qty) < Val(ir_qty) Then
ReDim Preserve arr(UBound(arr) + 1)
arr(UBound(arr)) = i
End If
i = i + 1
gr_qty = sheet.Cells(i, gr_col).Value
ir_qty = sheet.Cells(i, ir_col).Value
Wend
Case "ir"
While gr_qty <> ""
If Val(gr_qty) > Val(ir_qty) Then
ReDim Preserve arr(UBound(arr) + 1)
arr(UBound(arr)) = i
End If
i = i + 1
gr_qty = sheet.Cells(i, gr_col).Value
ir_qty = sheet.Cells(i, ir_col).Value
Wend
End Select
status = arr
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Status found"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
End Function

Function insert_employee(path As String, newsheet As Excel.Worksheet)
'open employee_info.xls
Set eApp = New Excel.Application
Set ebook = eApp.Workbooks.Open(path)
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "employee_info opened"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
eApp.Visible = False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "File not visible"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Set esheet = ebook.Worksheets(2)
'while loop
Dim ecell, newcell As String
Dim first_name, last_name As String
Dim last, first, user As Integer
Dim i, j As Integer
last = find_col("last name", esheet)
first = find_col("first name", esheet)
user = find_col("user name", esheet)
'initialize
newsheet.Cells(1, 1) = "Name"
i = 2
ecell = UCase(Trim(esheet.Cells(i, user).Value))
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Inserting names"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
While ecell <> ""
j = 2
newcell = UCase(Trim(newsheet.Cells(j, 2).Value))
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Naming: " & ecell
Form4.Text1.SelStart = Len(Form4.Text1.Text)
While newcell <> ""
If newcell = ecell Then
last_name = UCase(Trim(esheet.Cells(i, last).Value))
first_name = UCase(Trim(esheet.Cells(i, first).Value))
newsheet.Cells(j, 1).Value = last_name & " " & first_name
End If
j = j + 1
newcell = UCase(Trim(newsheet.Cells(j, 2).Value))
Wend
i = i + 1
ecell = UCase(Trim(esheet.Cells(i, user).Value))
Wend
'inserting N/A
Dim created As String
Dim k As Integer
k = 1
created = newsheet.Cells(k, 2).Value
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Filling N/A"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
While created <> ""
If newsheet.Cells(k, 1).Value = "" Then
newsheet.Cells(k, 1).Value = "N/A"
End If
k = k + 1
created = newsheet.Cells(k, 2).Value
Wend
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Finished inserting info"
'close employee_info.xls
ebook.Close savechanges:=False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "employee_info.xls closed"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
eApp.Quit
Set eApp = Nothing
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Pointer released"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
End Function

Function insert_row(opt As String, arr() As String, sheet As Excel.Worksheet, newsheet As Excel.Worksheet)
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Calling insert_row funtion"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Inserting open cases info"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
'Variables declaration
Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, status As Integer
'initializing
a = find_col("created", sheet)
b = find_col("purch.doc.", sheet)
c = find_col("item", sheet)
d = find_col("material", sheet)
e = find_col("short text", sheet)
f = find_col("wbs element", sheet)
g = find_col("document", sheet)
h = find_col("item", sheet)
i = find_col("order", sheet)
j = find_col("network", sheet)
k = find_col("vendor", sheet)
l = find_col("vendor name 1", sheet)
If opt = "gr" Then
m = find_col("open gr qty", sheet)
ElseIf opt = "ir" Then
m = find_col("open ir qty", sheet)
End If
status = UBound(arr)
'insertion loop
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Inserting case detail"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
For n = 0 To status
newsheet.Cells(n + 1, 2).Value = sheet.Cells(arr(n), a).Value
newsheet.Cells(n + 1, 3).Value = sheet.Cells(arr(n), b).Value
newsheet.Cells(n + 1, 4).Value = sheet.Cells(arr(n), c).Value
newsheet.Cells(n + 1, 5).Value = sheet.Cells(arr(n), d).Value
newsheet.Cells(n + 1, 6).Value = sheet.Cells(arr(n), e).Value
newsheet.Cells(n + 1, 7).Value = sheet.Cells(arr(n), f).Value
newsheet.Cells(n + 1, 8).Value = sheet.Cells(arr(n), g).Value
newsheet.Cells(n + 1, 9).Value = sheet.Cells(arr(n), h).Value
newsheet.Cells(n + 1, 10).Value = sheet.Cells(arr(n), i).Value
newsheet.Cells(n + 1, 11).Value = sheet.Cells(arr(n), j).Value
newsheet.Cells(n + 1, 12).Value = sheet.Cells(arr(n), k).Value
newsheet.Cells(n + 1, 13).Value = sheet.Cells(arr(n), l).Value
newsheet.Cells(n + 1, 14).Value = sheet.Cells(arr(n), m).Value
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Inserting case# " & n
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Next n
Call insert_employee(emp_path, newsheet)
'formating
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Formating"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
newsheet.Columns("A:N").AutoFit
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Autofit column width"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
End Function

Function run_gr(path As String)
'open file
Set grApp = New Excel.Application
Set grbook = grApp.Workbooks.Open(path)
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "open_gr.xls opened"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
grApp.Visible = False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "File not visible"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Set grsheet = grbook.Worksheets(1)
'calculate status
Dim arr() As String
Dim total_gr As Integer
arr() = status("gr", grsheet)
total_gr = UBound(arr)
'create new excel
Set newApp = New Excel.Application
Set newbook = newApp.Workbooks.Add
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Output excel created"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
newApp.Visible = False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "File not visible"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Set newsheet1 = newbook.Worksheets(1)
NewExcel = True
'insert info
Call insert_row("gr", arr, grsheet, newsheet1)
newsheet1.Name = "Open GR"
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Worksheet name changed"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
'close file
grbook.Close savechanges:=False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "open_gr.xls closed"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
grApp.Quit
Set grApp = Nothing
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "pointer released"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Label5.Caption = total_gr
End Function

Function run_ir(path As String)
'open file
Set irApp = New Excel.Application
Set irbook = irApp.Workbooks.Open(path)
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "open_ir.xls opened"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
irApp.Visible = False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "File not visible"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Set irsheet = irbook.Worksheets(1)
'calculate status
Dim arr() As String
Dim total_ir As Integer
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Calculating open status..."
Form4.Text1.SelStart = Len(Form4.Text1.Text)
arr() = status("ir", irsheet)
total_ir = UBound(arr)
'create new excel
Set newApp = New Excel.Application
Set newbook = newApp.Workbooks.Add
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Output excel created"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
newApp.Visible = False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "File not visible"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Set newsheet1 = newbook.Worksheets(1)
NewExcel = True
'insert info
Call insert_row("ir", arr, irsheet, newsheet1)
newsheet1.Name = "Open IR"
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Worksheet name changed"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
'close file
irbook.Close savechanges:=False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "open_ir.xls closed"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
irApp.Quit
Set irApp = Nothing
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "pointer released"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Label6.Caption = total_ir
End Function

Function run_both(grpath As String, irpath As String)
'calculating ir
Call run_ir(irpath)
'calculating gr
'open file
Set grApp = New Excel.Application
Set grbook = grApp.Workbooks.Open(grpath)
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "open_gr.xls opened"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
grApp.Visible = False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "File not visible"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Set grsheet = grbook.Worksheets(1)
'calculate status
Dim arr() As String
Dim total_gr As Integer
arr() = status("gr", grsheet)
total_gr = UBound(arr)
'add new sheet
Set newsheet2 = newApp.Application.Worksheets.Add
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "new worksheet added"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
'insert info
Call insert_row("gr", arr, grsheet, newsheet2)
newsheet2.Name = "Open GR"
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "Worksheet name changed"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
'close file
grbook.Close savechanges:=False
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "open_gr.xls closed"
grApp.Quit
Set grApp = Nothing
Form4.Text1.Text = Form4.Text1.Text & vbCrLf & "pointer released"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Label5.Caption = total_gr
End Function

Private Sub delay(sec As Single)
Dim tmrEnd
tmrEnd = Timer() + sec
Do While Timer() < tmrEnd
DoEvents
Loop
End Sub

Private Sub Form_Load()
Me.Hide
frmSplash.Show
NewExcel = False
Saved = False
gr_path = "C:\Users\Vivien Fan\My Documents\^~ ^ Vi Vi Ann\SBT\Commercial\open_GR.xls"
ir_path = "C:\Users\Vivien Fan\My Documents\^~ ^ Vi Vi Ann\SBT\Commercial\open_IR.xls"
emp_path = "C:\Users\Vivien Fan\My Documents\^~ ^ Vi Vi Ann\SBT\Commercial\employee_info.xls"
delay (0.5)
Unload frmSplash
Me.Show
End Sub

Private Sub command1_click()
Select Case search()
Case "gi"
Form4.Show
Form4.Text1.Text = "Calculating open gr & ir cases"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Call run_both(gr_path, ir_path)
Unload Form4
Case "g"
Form4.Show
Form4.Text1.Text = "Calculating open gr cases"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Call run_gr(gr_path)
Unload Form4
Case "i"
Form4.Show
Form4.Text1.Text = "Calculating open ir cases"
Form4.Text1.SelStart = Len(Form4.Text1.Text)
Call run_ir(ir_path)
Unload Form4
Case ""
MsgBox "Error: Please select at least one option", vbOKOnly, "Error"
End Select
End Sub

Private Sub Command2_Click()
If NewExcel Then
newApp.Visible = True
NewExcel = False
Else
MsgBox "Error: Object does not exist", vbOKOnly, "Error"
End If
End Sub

Private Sub Command3_Click()
If NewExcel Then
Form5.Show
Else
MsgBox "Error: Object does not exist", vbOKOnly, "error"
End If
End Sub

Private Sub Command4_Click()
If NewExcel Then
If MsgBox("You haven't saved the result excel. Do you want to exit anyway?", vbYesNo, "Open Cases Searching Engine") = vbYes Then
'user decided to exist the application without saving the new excel
newApp.DisplayAlerts = False
newbook.Close savechanges:=False
newApp.DisplayAlerts = True
newApp.Quit
Set newApp = Nothing
Unload Me
Else
'user decided to save it, then we open the excel file for him/her
newApp.Visible = True
NewExcel = False
End If
Else
'we dont have a new excel file
'therefore there is no pointer needed to be freed
Unload Me
End If
End Sub

Private Sub GRLocation_Click()
Form2.Show
End Sub

Private Sub IRLocation_Click()
Form3.Show
End Sub

Private Sub emp_info_Click()
Form6.Show
End Sub

Private Sub Exit_Click()
If NewExcel Then
'we have a new excel file
If Saved Then 'we have already saved it or opened it
Unload Me
Else 'we havn't saved it or opened it yet
If MsgBox("You haven't saved the result excel. Do you want to exit anyway?", vbYesNo, "Open Cases Searching Engine") = vbYes Then
'user decided to exist the application without saving the new excel
newApp.DisplayAlerts = False
newbook.Close
newApp.DisplayAlerts = True
newApp.Quit
Set newApp = Nothing
Unload Me
Else
'user decided to save it, then we open the excel file for him/her
newApp.Visible = True
End If
End If
Else
'we dont have a new excel file
'therefore there is no pointer needed to be freed
Unload Me
End If
End Sub
