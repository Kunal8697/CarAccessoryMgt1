VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEmp 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   105
   ClientTop       =   720
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9270
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   938
      TabIndex        =   16
      Top             =   720
      Width           =   7305
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2625
         Left            =   360
         TabIndex        =   17
         Top             =   180
         Width           =   6495
         Begin MSComCtl2.DTPicker DTPJDate 
            Height          =   375
            Left            =   2160
            TabIndex        =   2
            ToolTipText     =   "Select Joining Date"
            Top             =   1440
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin VB.TextBox TxtBpay 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2160
            TabIndex        =   3
            ToolTipText     =   "Enter Basic Pay"
            Top             =   2070
            Width           =   3015
         End
         Begin VB.TextBox TxtECode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   2160
            TabIndex        =   0
            ToolTipText     =   "Enter Employee Code"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtEName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2160
            TabIndex        =   1
            Top             =   900
            Width           =   3135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Basic Pay"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   900
            TabIndex        =   21
            Top             =   2070
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Joining Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   615
            TabIndex        =   20
            Top             =   1530
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Emp Code"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   855
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Emp Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   780
            TabIndex        =   18
            Top             =   900
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1635
      Left            =   938
      TabIndex        =   4
      Top             =   4050
      Width           =   7395
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   360
         TabIndex        =   5
         Top             =   90
         Width           =   6765
         Begin VB.CommandButton CmdDelete 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4050
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton CmdAdd 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   270
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton CmdSave 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1530
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton CmdFirst 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&First"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   270
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdNext 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Next"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1530
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdPrevious 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Previous"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4050
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdLast 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Last"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdExit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1275
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3570
      TabIndex        =   15
      Top             =   180
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2610
      Picture         =   "FrmEmp.frx":0000
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "FrmEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmp As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset
Dim q As String
Private Sub CmdAdd_Click()
Call ClearText(Me)
Frame2.Enabled = True
Dim a As Integer
a = rsCode(2)
TxtECode.Text = a + 1
rsCode.Fields(2) = Val(TxtECode.Text)
rsCode.Update
rsEmp.AddNew
End Sub
Private Sub CmdDelete_Click()
rsEmp.Delete
Call ClearText(Me)
rsEmp.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsEmp.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsEmp.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsEmp.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsEmp.MoveNext
    If rsEmp.EOF = False Then
        Call RecordsetToText
    Else
        rsEmp.MoveFirst
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsEmp.MovePrevious
    If rsEmp.BOF = False Then
        Call RecordsetToText
    Else
        rsEmp.MoveLast
        Call RecordsetToText
    End If
End Sub


Private Sub CmdSave_Click()
Frame2.Enabled = True
Call TextToRecordset
rsEmp.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from EmpMaster"
rsEmp.CursorLocation = adUseClient
rsEmp.Open q, cn, adOpenKeyset, adLockOptimistic
'************
rsCode.CursorLocation = adUseClient
rsCode.Open "CodeSet", cn, adOpenKeyset, adLockOptimistic
End Sub
Public Function TextToRecordset()
rsEmp.Fields(0) = Val(TxtECode.Text)
rsEmp.Fields(1) = IntoStr(Trim(TxtEName.Text))
rsEmp.Fields(2) = DTPJDate.Value
rsEmp.Fields(3) = Val(TxtBpay.Text)
End Function
Public Function RecordsetToText()
 TxtECode.Text = rsEmp.Fields(0)
 TxtEName.Text = rsEmp.Fields(1)
 DTPJDate.Value = rsEmp.Fields(2)
 TxtBpay.Text = rsEmp.Fields(3)
 End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsEmp = Nothing
Set rsCode = Nothing

End Sub
Private Sub TxtEName_KeyPress(KeyAscii As Integer)
Call CheckName(KeyAscii)
End Sub
