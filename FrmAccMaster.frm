VERSION 5.00
Begin VB.Form FrmAccMaster 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8640
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   581
      TabIndex        =   11
      Top             =   4410
      Width           =   7485
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   270
         TabIndex        =   12
         Top             =   90
         Width           =   7035
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
            Left            =   5430
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   720
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
            TabIndex        =   17
            Top             =   720
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
            TabIndex        =   16
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   270
            Width           =   1275
         End
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
            TabIndex        =   13
            Top             =   270
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   574
      TabIndex        =   5
      Top             =   1080
      Width           =   7395
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2715
         Left            =   630
         TabIndex        =   6
         Top             =   180
         Width           =   6225
         Begin VB.TextBox TxtPrice 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2400
            TabIndex        =   4
            ToolTipText     =   "Enter Price Of Accessories"
            Top             =   1980
            Width           =   1425
         End
         Begin VB.TextBox TxtSrNo 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3960
            TabIndex        =   1
            ToolTipText     =   "Enter Series Of Accessories"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtAccPeriod 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2430
            TabIndex        =   3
            ToolTipText     =   "Enter Warrenty Period"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox TxtAccName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2430
            TabIndex        =   2
            ToolTipText     =   "Enter Name Of Accessories"
            Top             =   900
            Width           =   3135
         End
         Begin VB.TextBox TxtAccCode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   2430
            TabIndex        =   0
            ToolTipText     =   "Enter Accessory Code"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "In Day's"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3870
            TabIndex        =   23
            Top             =   1530
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Price"
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
            Left            =   1710
            TabIndex        =   22
            Top             =   1980
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Warranty Period"
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
            Left            =   480
            TabIndex        =   10
            Top             =   1440
            Width           =   1770
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Accessories Name"
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
            Left            =   255
            TabIndex        =   9
            Top             =   900
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Accessories Code"
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
            Left            =   330
            TabIndex        =   8
            Top             =   360
            Width           =   1920
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2160
      Picture         =   "FrmAccMaster.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Accessories Setting"
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
      Left            =   3135
      TabIndex        =   7
      Top             =   540
      Width           =   2535
   End
End
Attribute VB_Name = "FrmAccMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAMast As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset
Dim q As String
Private Sub CmdAdd_Click()
Frame2.Enabled = True
Call ClearText(Me)
Dim a As Integer
a = rsCode(0)
TxtAccCode.Text = a + 1
rsCode.Fields(0) = Val(TxtAccCode.Text)
rsCode.Update
rsAMast.AddNew
End Sub
Private Sub CmdDelete_Click()
rsAMast.Delete
Call ClearText(Me)
rsAMast.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsAMast.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsAMast.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsAMast.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsAMast.MoveNext
    If rsAMast.EOF = False Then
        Call RecordsetToText
    Else
        rsAMast.MoveFirst
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsAMast.MovePrevious
    If rsAMast.BOF = False Then
        Call RecordsetToText
    Else
        rsAMast.MoveLast
        Call RecordsetToText
    End If
End Sub


Private Sub CmdSave_Click()
Frame2.Enabled = True
Call TextToRecordset
rsAMast.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from AccessaryMaster"
rsAMast.CursorLocation = adUseClient
rsAMast.Open q, cn, adOpenKeyset, adLockOptimistic
'**************
rsCode.CursorLocation = adUseClient
rsCode.Open "CodeSet", cn, adOpenKeyset, adLockOptimistic
End Sub
Public Function TextToRecordset()
rsAMast.Fields(0) = Val(TxtAccCode.Text)
rsAMast.Fields(1) = IntoStr(Trim(TxtAccName.Text))
rsAMast.Fields(2) = Val(TxtAccPeriod.Text)
rsAMast.Fields(3) = IntoStr(Trim(TxtSrNo.Text))
rsAMast.Fields(4) = IntoStr(Val(TxtPrice.Text))
End Function
Public Function RecordsetToText()
 TxtAccCode.Text = rsAMast.Fields(0)
 TxtAccName.Text = rsAMast.Fields(1)
 TxtAccPeriod.Text = rsAMast.Fields(2)
 TxtSrNo.Text = rsAMast.Fields(3)
 TxtPrice.Text = rsAMast.Fields(4)
 End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsAMast = Nothing
Set rsCode = Nothing
End Sub
Private Sub TxtAccName_KeyPress(KeyAscii As Integer)
Call CheckName(KeyAscii)
End Sub
Private Sub TxtAccName_LostFocus()
TxtAccName.Text = UCase(TxtAccName.Text)
End Sub
Private Sub TxtSrNo_LostFocus()
TxtSrNo.Text = UCase(TxtSrNo.Text)
End Sub
