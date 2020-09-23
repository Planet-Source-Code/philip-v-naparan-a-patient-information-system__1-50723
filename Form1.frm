VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Information System"
   ClientHeight    =   3450
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6780
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   6855
      TabIndex        =   34
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   6855
         TabIndex        =   35
         Top             =   15
         Width           =   6855
      End
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Height          =   615
      Left            =   3720
      Picture         =   "Form1.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "Form1.frx":1054
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   840
      Picture         =   "Form1.frx":1D1E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   1560
      Picture         =   "Form1.frx":29E8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   2280
      Picture         =   "Form1.frx":36B2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   3000
      Picture         =   "Form1.frx":437C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   800
      Left            =   4680
      TabIndex        =   12
      Top             =   2520
      Width           =   2000
      Begin VB.CommandButton Command14 
         Height          =   300
         Left            =   1470
         Picture         =   "Form1.frx":5046
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   500
      End
      Begin VB.CommandButton Command13 
         Height          =   300
         Left            =   990
         Picture         =   "Form1.frx":5190
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   500
      End
      Begin VB.CommandButton Command12 
         Height          =   300
         Left            =   510
         Picture         =   "Form1.frx":52DA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   500
      End
      Begin VB.CommandButton Command11 
         Height          =   300
         Left            =   15
         Picture         =   "Form1.frx":5424
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   500
      End
      Begin VB.TextBox txtcount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   30
         TabIndex        =   19
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblmax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1090
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "  of"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   3720
      Picture         =   "Form1.frx":556E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Height          =   615
      Left            =   3000
      Picture         =   "Form1.frx":6238
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Delete Flag"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Room #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   32
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Admit Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Patient Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Patient #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   27
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1545
      TabIndex        =   26
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   24
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   3000
      TabIndex        =   22
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   3000
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   36
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuNF 
         Caption         =   "&New File"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuO 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSFA 
         Caption         =   "&Save File As..."
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuE 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Dim Add As Boolean
Dim cur_num_holder As Integer
Dim tmp_patient_holder As String
Private Sub Command1_Click()
If Add = False Then
    Call unlock_txt
    Edit = False
    Add = True
    Call hide_controls
    Call clear_txt
    Command1.Left = Command4.Left
    Label5.Caption = "Add"
    Label5.Left = Label10.Left
    Label13.Visible = False
    Text6.Visible = False
    mnuNF.Enabled = False
    mnuO.Enabled = False
    mnuSFA.Enabled = False
Else
    If clecked_txt = False Then Exit Sub
    If record_exist(Text1.Text, total_num) = True And sDelFlag = "Y" Then
        patient(sDelFlag_num, 1) = Text1.Text
        patient(sDelFlag_num, 2) = Text2.Text
        patient(sDelFlag_num, 3) = Text3.Text
        patient(sDelFlag_num, 4) = Text4.Text
        patient(sDelFlag_num, 5) = Text5.Text
        patient(sDelFlag_num, 6) = "N"
        Call write_file(total_num)
        MsgBox "The deleted Patient #" & Text1.Text & " has been successfully change.", vbInformation, "Patient Info"
        Command5_Click
        curr_num = sDelFlag_num
        Command6_Click
    ElseIf record_exist(Text1.Text, total_num) = True And sDelFlag = "N" Then
        MsgBox "Patient #" & Text1.Text & " was already exist. Please change it !.", vbExclamation, "Patient Info"
        Text1.SetFocus
    Else
        If curr_num = 0 Then
            Open file_path For Output As #4
        Else
            Open file_path For Append As #4
        End If
            Print #4, Text1.Text & vbTab & Text2.Text & vbTab & Text3.Text & vbTab & Text4.Text & vbTab & Text5.Text & vbTab & "N"
        Close #4
        MsgBox "New record has been successfully added.", vbInformation, "Patient Info"
        Command5_Click
        curr_num = total_num
        Command6_Click
    End If
End If
End Sub

Private Sub Command10_Click()
'[ Terminate application ]
Unload Me
End Sub

Private Sub Command11_Click()
If curr_num > 0 Then curr_num = 1
Call navigate(1, total_num)
End Sub

Private Sub Command12_Click()
If curr_num > 1 Then
    curr_num = curr_num - 1
    Call navigate(curr_num, total_num)
ElseIf curr_num = 1 Then
    Call navigate(1, total_num)
End If
End Sub

Private Sub Command13_Click()
If curr_num < total_num Then
    curr_num = curr_num + 1
    Call navigate(curr_num, total_num)
ElseIf curr_num = total_num Then
    Call navigate(total_num, total_num)
End If
End Sub

Private Sub Command14_Click()
If curr_num > 0 Then curr_num = total_num
Call navigate(total_num, total_num)
End Sub

Private Sub Command2_Click()
If total_num = 0 Then MsgBox "No record to edit.", vbExclamation, "Patient Info": Exit Sub
If patient(curr_num, 6) = "Y" Then MsgBox "This record cannot be edit because it was deleted." & vbCrLf & "Please add a record with the same record number w/ the deleted record to replace it.", vbInformation, "Patient Info": Exit Sub
If Edit = False Then
    Call unlock_txt
    Edit = True
    Call hide_controls
    Command2.Left = Command4.Left
    Label6.Caption = "Save"
    Label6.Left = Label10.Left
    cur_num_holder = curr_num
    tmp_patient_holder = LCase(Text1.Text)
    mnuNF.Enabled = False
    mnuO.Enabled = False
    mnuSFA.Enabled = False
ElseIf record_exist(Text1.Text, total_num) = True And sDelFlag = "N" And LCase(Text1.Text) <> tmp_patient_holder Then
        MsgBox "Patient #" & Text1.Text & " was already exist. Please change it !.", vbExclamation, "Patient Info"
        Text1.SetFocus
Else
    If clecked_txt = False Then Exit Sub
    patient(curr_num, 1) = Text1.Text
    patient(curr_num, 2) = Text2.Text
    patient(curr_num, 3) = Text3.Text
    patient(curr_num, 4) = Text4.Text
    patient(curr_num, 5) = Text5.Text
    patient(curr_num, 6) = Text6.Text
    Call write_file(total_num)
    MsgBox "Changes has been successfully saved.", vbInformation, "Patient Info"
    Command5_Click
    curr_num = cur_num_holder
    Command6_Click
End If
End Sub

Private Sub Command3_Click()
If total_num = 0 Then MsgBox "No record to search.", vbExclamation, "Patient Info": Exit Sub
Dim src As String
src = InputBox("Enter Patient # to search.", "Search Patient", "Type Here !")
If src = "" Or src = "Type Here !" Then
    MsgBox "Search has not been done!", vbExclamation, "Patient Info"
Else
    Call search_record(src, total_num, curr_num)
    Call navigate(curr_num, total_num)
End If
End Sub

Private Sub Command4_Click()
If total_num = 0 Then MsgBox "No record to delete.", vbExclamation, "Patient Info": Exit Sub
If patient(curr_num, 6) = "Y" Then MsgBox "Record already deleted.", vbInformation, "Patient Info": Exit Sub
Dim reply As Integer
reply = MsgBox("Are you sure you want to delete the current record?", vbCritical + vbYesNo, "Confirm Delete")
If reply = vbYes Then
    patient(curr_num, 1) = Text1.Text
    patient(curr_num, 2) = Text2.Text
    patient(curr_num, 3) = Text3.Text
    patient(curr_num, 4) = Text4.Text
    patient(curr_num, 5) = Text5.Text
    patient(curr_num, 6) = "Y"
    Call write_file(total_num)
    cur_num_holder = curr_num
    Command5_Click
    curr_num = cur_num_holder
    Call navigate(curr_num, total_num)
    MsgBox "Record has been successfully deleted.", vbInformation, "Patient Info"
End If
End Sub

Private Sub Command5_Click()
Call clear_txt
total_num = file_total
Call load_file
Call navigate(curr_num, total_num)
End Sub

Private Sub Command6_Click()
Call lock_txt
Call show_controls
If Edit = True Then
    Command2.Left = 840
    Label6.Left = 840
    Label6.Caption = "Edit"
    tmp_patient_holder = ""
    Edit = False
    Command2.SetFocus
Else
    If Add = True Then
        Command1.Left = 120
        Label5.Left = 120
        Label5.Caption = "Add"
        Label13.Visible = True
        Text6.Visible = True
        Add = False
    End If
    Command1.SetFocus
End If
mnuNF.Enabled = True
mnuO.Enabled = True
mnuSFA.Enabled = True
Call navigate(curr_num, total_num)
End Sub

Private Sub Command7_Click()
Call clear_txt
End Sub
Private Sub Form_Load()
Call lock_control
Call lock_txt
Add = False
Edit = False
CommonDialog1.Filter = "Patient File (*.p)|*.p"

Me.Show
frmSplash.Show vbModal

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reply As Integer
reply = MsgBox("Are you sure you want close this application?", vbExclamation + vbYesNo, "Patient Info")
If reply = vbNo Then
    Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuE_Click()
Unload Me
End Sub

Private Sub mnuNF_Click()
Me.MousePointer = vbHourglass
    CommonDialog1.DialogTitle = "New Patient File"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Call clear_txt
    curr_num = 0
    total_num = 0
    file_path = CommonDialog1.FileName
    Open file_path For Output As #5
    Close #5
    Call unlock_control
    Call lock_txt
    ReDim patient(0, 0)
    Call navigate(curr_num, total_num)
    Command1.SetFocus
    CommonDialog1.FileName = ""
Me.MousePointer = vbDefault
End Sub

Private Sub mnuO_Click()
Me.MousePointer = vbHourglass
    CommonDialog1.DialogTitle = "Open Patient File"
    CommonDialog1.CancelError = False
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    curr_num = 0
    Call clear_txt
    file_path = CommonDialog1.FileName
    total_num = file_total
    Call unlock_control
    Call lock_txt
    Call load_file
    Call navigate(curr_num, total_num)
    Command1.SetFocus
    CommonDialog1.FileName = ""
Me.MousePointer = vbDefault
End Sub
Private Sub navigate(ByVal sNum As Integer, ByVal sTotal As Integer)
If curr_num = 0 Then txtcount.Text = 0: lblmax = 0: Exit Sub

Text1.Text = patient(sNum, 1)
Text2.Text = patient(sNum, 2)
Text3.Text = patient(sNum, 3)
Text4.Text = patient(sNum, 4)
Text5.Text = patient(sNum, 5)
Text6.Text = patient(sNum, 6)

txtcount.Text = sNum
lblmax = sTotal

End Sub

Private Sub mnuSFA_Click()
Me.MousePointer = vbHourglass
    CommonDialog1.DialogTitle = "Save Patient As.."
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    file_path = CommonDialog1.FileName
    Call write_file(total_num)
    Call navigate(curr_num, total_num)
    Command1.SetFocus
    CommonDialog1.FileName = ""
Me.MousePointer = vbDefault

End Sub

Private Sub txtcount_Change()
If Not Val(txtcount) < 1 And Not Val(txtcount) > total_num Then Call navigate(Val(txtcount), total_num)
End Sub

Private Sub txtcount_GotFocus()
'[ Highlight the text ]
txtcount.SelStart = 0
txtcount.SelLength = Len(txtcount.Text)
End Sub

Private Sub txtcount_KeyPress(KeyAscii As Integer)
'[ Disable character input ]
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub clear_txt()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub
Private Sub lock_txt()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
End Sub
Private Sub unlock_txt()
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
End Sub
Private Sub hide_controls()
If Edit = True Then
    If Add = False Then
        Command1.Visible = False
        Label5.Visible = False
    End If
Else
    Command2.Visible = False
    Label6.Visible = False
End If
Command3.Visible = False
    Label11.Visible = False
Command4.Visible = False
    Label10.Visible = False
Command5.Visible = False
    Label3.Visible = False
Command10.Visible = False
    Label8.Visible = False
Frame1.Visible = False

Command6.Visible = True
Label14.Visible = True
Command7.Visible = True
Label15.Visible = True
End Sub
Private Sub show_controls()
If Edit = True Then
    If Add = False Then
        Command1.Visible = True
        Label5.Visible = True
    End If
Else
    Command2.Visible = True
    Label6.Visible = True
End If
Command3.Visible = True
    Label11.Visible = True
Command4.Visible = True
    Label10.Visible = True
Command5.Visible = True
    Label3.Visible = True
Command10.Visible = True
    Label8.Visible = True
Frame1.Visible = True

Command6.Visible = False
Label14.Visible = False
Command7.Visible = False
Label15.Visible = False
End Sub
Private Function clecked_txt() As Boolean
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Some fields is/are empty.Please check it.", vbExclamation, "Patient Info"
    clecked_txt = False
Else
    clecked_txt = True
End If
End Function
Public Sub lock_control()
Frame1.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command10.Enabled = False

Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False
Command14.Enabled = False

Text1.BackColor = &H8000000F
Text2.BackColor = &H8000000F
Text3.BackColor = &H8000000F
Text4.BackColor = &H8000000F
Text5.BackColor = &H8000000F
Text6.BackColor = &H8000000F
End Sub
Public Sub unlock_control()
Frame1.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command10.Enabled = True

Command11.Enabled = True
Command12.Enabled = True
Command13.Enabled = True
Command14.Enabled = True

Text1.BackColor = &H80000005
Text2.BackColor = &H80000005
Text3.BackColor = &H80000005
Text4.BackColor = &H80000005
Text5.BackColor = &H80000005
Text6.BackColor = &H80000005
End Sub
