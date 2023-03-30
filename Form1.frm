VERSION 5.00
Begin VB.Form FrmAccMaster 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   450
      TabIndex        =   9
      Top             =   4410
      Width           =   7395
      Begin VB.Frame Frame6 
         BackColor       =   &H00EE8EDD&
         Height          =   1275
         Left            =   180
         TabIndex        =   10
         Top             =   90
         Width           =   7035
         Begin VB.CommandButton CmdExit 
            BackColor       =   &H8000000A&
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
            Height          =   375
            Left            =   5310
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "P&rint"
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
            Left            =   5310
            TabIndex        =   19
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton CmdLast 
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
            TabIndex        =   18
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdPrevious 
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
            TabIndex        =   17
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdNext 
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
            TabIndex        =   16
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdFirst 
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
            TabIndex        =   15
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdSave 
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
            TabIndex        =   14
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton CmdAdd 
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
            TabIndex        =   13
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H8000000A&
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
            TabIndex        =   12
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton CmdDelete 
            BackColor       =   &H8000000A&
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
            TabIndex        =   11
            Top             =   270
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   7395
      Begin VB.Frame Frame2 
         BackColor       =   &H00EE8EDD&
         Height          =   2715
         Left            =   630
         TabIndex        =   1
         Top             =   180
         Width           =   6225
         Begin VB.TextBox TxtPrice 
            Height          =   375
            Left            =   2700
            TabIndex        =   22
            Top             =   2070
            Width           =   1425
         End
         Begin VB.TextBox TxtSrNo 
            Height          =   375
            Left            =   4320
            TabIndex        =   21
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtAccPeriod 
            Height          =   375
            Left            =   2700
            TabIndex        =   6
            Top             =   1530
            Width           =   3135
         End
         Begin VB.TextBox TxtAccName 
            Height          =   375
            Left            =   2700
            TabIndex        =   5
            Top             =   900
            Width           =   3135
         End
         Begin VB.TextBox TxtAccCode 
            Height          =   375
            Left            =   2790
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1620
            TabIndex        =   23
            Top             =   2070
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Warrenty Period"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   450
            TabIndex        =   8
            Top             =   1440
            Width           =   1665
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Accessories Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   450
            TabIndex        =   7
            Top             =   900
            Width           =   1860
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "A ccessories Code"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   450
            TabIndex        =   3
            Top             =   360
            Width           =   1845
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1350
      Picture         =   "Form1.frx":0000
      Top             =   450
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00EE8EDD&
      Caption         =   "Accessories Setting"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2340
      TabIndex        =   2
      Top             =   540
      Width           =   2370
   End
End
Attribute VB_Name = "FrmAccMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsAMast As New ADODB.Recordset
Dim q As String

Private Sub CmdAdd_Click()
rsAMast.AddNew
End Sub

Private Sub CmdDelete_Click()
rsAMast.Delete
Call ClearText(Me)
rsAMast.MoveFirst
Call RecordsetToText

End Sub

Private Sub CmdEdit_Click()
Call TextToRecordset
rsAMast.Update
End Sub

Private Sub CmdExit_Click()
Unload Me


End Sub

Private Sub CmdFirst_Click()
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsAMast.MoveFirst
Call RecordsetToText

End Sub

Private Sub CmdLast_Click()
CmdNext.Enabled = False
rsAMast.MoveLast
Call RecordsetToText

End Sub

Private Sub CmdNext_Click()
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
Call TextToRecordset
rsAMast.Update
Call ClearText(Me)
End Sub


Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from AccessaryMaster"
rsAMast.CursorLocation = adUseClient
rsAMast.Open q, cn, adOpenKeyset, adLockOptimistic

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

End Sub

Private Sub TxtSrNo_LostFocus()
TxtSrNo.Text = UCase(TxtSrNo.Text)
End Sub
