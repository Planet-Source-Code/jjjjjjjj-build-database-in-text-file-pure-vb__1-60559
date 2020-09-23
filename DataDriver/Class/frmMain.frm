VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "DataBaseDriver - In Pure VB [ By, Jim Jose ]"
   ClientHeight    =   5805
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Find"
      Height          =   1095
      Left            =   2040
      TabIndex        =   34
      Top             =   4200
      Width           =   4335
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   3480
         TabIndex        =   38
         Text            =   "-1"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "FindText"
         Height          =   375
         Left            =   1800
         TabIndex        =   37
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtFind 
         Height          =   360
         Left            =   120
         TabIndex        =   36
         Text            =   "Emp Name 333"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtStart 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   3480
         TabIndex        =   35
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Field"
         Height          =   240
         Left            =   2880
         TabIndex        =   40
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Start"
         Height          =   240
         Left            =   2880
         TabIndex        =   39
         Top             =   720
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C9F1FC&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   8685
      TabIndex        =   32
      Top             =   0
      Width           =   8745
      Begin VB.Label Label14 
         BackColor       =   &H00C9F1FC&
         Caption         =   $"frmMain.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   8655
      End
   End
   Begin VB.PictureBox picTest 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   6570
      ScaleHeight     =   4425
      ScaleWidth      =   2145
      TabIndex        =   16
      Top             =   975
      Width           =   2175
      Begin VB.CommandButton cmdTest 
         Caption         =   "Save +1,000 Records"
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lbSize 
         AutoSize        =   -1  'True
         Caption         =   "N/a"
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Test Purpose"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   27
         Top             =   120
         Width           =   1380
      End
   End
   Begin VB.PictureBox picOperations 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   1815
      TabIndex        =   11
      Top             =   975
      Width           =   1845
      Begin VB.CommandButton cmd_DelAll 
         Caption         =   "Delete All"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open DataBase>"
         Height          =   855
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save New"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Now"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Now"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operations"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8685
      TabIndex        =   9
      Top             =   5430
      Width           =   8745
      Begin VB.Label lbCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   8280
         TabIndex        =   29
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lbStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The File is not opened. Please 'OPEN DATABASE>'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   3600
      End
   End
   Begin VB.TextBox txtBox 
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtBox 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtBox 
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtBox 
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   ">"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdMovePrev 
      Caption         =   "<"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Size [20]"
      Height          =   240
      Left            =   5640
      TabIndex        =   26
      Top             =   3120
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Size [30]"
      Height          =   240
      Left            =   5640
      TabIndex        =   25
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Size [20]"
      Height          =   240
      Left            =   5640
      TabIndex        =   24
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Size [30]"
      Height          =   240
      Left            =   5640
      TabIndex        =   23
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label lbField 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/a"
      Height          =   240
      Left            =   2040
      TabIndex        =   22
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Salary"
      Height          =   240
      Left            =   2040
      TabIndex        =   21
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Desgn"
      Height          =   240
      Left            =   2040
      TabIndex        =   20
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Age"
      Height          =   240
      Left            =   2040
      TabIndex        =   19
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   240
      Left            =   2040
      TabIndex        =   18
      Top             =   1680
      Width           =   555
   End
   Begin VB.Label lbPos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/a"
      Height          =   240
      Left            =   3120
      TabIndex        =   6
      Top             =   3840
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents dBase As DataBaseDriver
Attribute dBase.VB_VarHelpID = -1

Private Sub cmd_DelAll_Click()
    dBase.DeleteAll
End Sub

Private Sub Form_Load()
    Set dBase = New DataBaseDriver
    cmdOpen_Click
    If dBase.RecordCount > 0 Then dBase.MoveFirst
End Sub

Private Sub cmdDelete_Click()
    dBase.Delete
    dBase.MoveFirst
End Sub

Private Sub cmdFirst_Click()
    dBase.MoveFirst
End Sub

Private Sub cmdLast_Click()
    dBase.MoveLast
End Sub

Private Sub cmdMoveNext_Click()
    dBase.MoveNext
End Sub

Private Sub cmdMovePrev_Click()
    dBase.MovePrevious
End Sub

Private Sub cmdOpen_Click()
    dBase.OpenFile App.Path & "\DataFile.txt"
    If dBase.RecordCount = 0 Then MsgBox "The opened file contains no data. Please run the test save of 1,000 records"
    lbStatus = "DataBase opened..."
End Sub

Private Sub cmdSave_Click()
    dBase.Save
End Sub

Private Sub cmdUpdate_Click()
    dBase.Update
End Sub

Private Sub cmdFind_Click()
Dim t As Double
t = Timer
    txtStart = dBase.FindText(txtFind, txtStart, txtField)
    lbStatus = "Time taken to search " & Timer - t & " Seconds"
    If txtStart = -1 Then MsgBox "Data not found!": txtStart = 0: Exit Sub
dBase.AbsolutePosition = txtStart
txtStart = txtStart + 1
End Sub

Private Sub cmdTest_Click()
Dim t As Double
Dim x As Long
lbStatus = "Starts saving data. Please Wait.."
DoEvents

t = Timer
For x = dBase.RecordCount To dBase.RecordCount + 1000
    txtBox(0) = "Emp Name " & x
    txtBox(1) = "AGE " & Int(Rnd * 60)
    txtBox(2) = "Designation " & x
    txtBox(3) = "Salary " & Int(Rnd * 10000)
    dBase.Save
Next x
lbStatus = "Completed! Time taken to save +1,000 records.. " & Timer - t & " Seconds"
dBase.MoveLast
End Sub

Private Sub txtFind_Change()
    txtStart = 0
End Sub

'The Class Events
'These are needed to transfer data b/w the field and Class

Private Sub dBase_RecordChanged()
    txtBox(0) = dBase.Field(0)
    txtBox(1) = dBase.Field(1)
    txtBox(2) = dBase.Field(2)
    txtBox(3) = dBase.Field(3)
    lbPos = "Position " & dBase.AbsolutePosition
    lbCount = "RecordCount > " & dBase.RecordCount
    lbField = "FieldCount  > " & dBase.FieldCount
    lbSize = "File Size " & Round(FileLen(App.Path & "\DataFile.txt") / 1024, 2) & " Kb"
End Sub

Private Sub dBase_BeforeSave()
    With dBase
        .Field(0) = txtBox(0)
        .Field(1) = txtBox(1)
        .Field(2) = txtBox(2)
        .Field(3) = txtBox(3)
    End With
End Sub

Private Sub dBase_BeforeUpdate()
    dBase_BeforeSave
End Sub
