VERSION 5.00
Begin VB.Form FrmSMast 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8070
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   7485
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1365
         Left            =   360
         TabIndex        =   9
         Top             =   180
         Width           =   6855
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            Left            =   5430
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2715
      Left            =   270
      TabIndex        =   4
      Top             =   810
      Width           =   7485
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2265
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   6225
         Begin VB.TextBox TxtSprice 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2340
            TabIndex        =   2
            ToolTipText     =   "Enter Price Of Service"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox TxtSCode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   2340
            TabIndex        =   0
            ToolTipText     =   "Enter Service Code"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtSName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2340
            TabIndex        =   1
            ToolTipText     =   "Enter Name Of Service"
            Top             =   900
            Width           =   3135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Price Of Service"
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
            Left            =   525
            TabIndex        =   19
            Top             =   1440
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Service Code"
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
            Left            =   810
            TabIndex        =   7
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name Of Service "
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
            Left            =   375
            TabIndex        =   6
            Top             =   900
            Width           =   1860
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Service Details"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   150
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   690
      Picture         =   "FrmSMast.frx":0000
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "FrmSMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsService As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset

Dim q As String
Private Sub CmdAdd_Click()
Call ClearText(Me)
Frame2.Enabled = True
Dim a As Integer
a = rsCode(1)
TxtSCode.Text = a + 1
rsCode.Fields(1) = Val(TxtSCode.Text)
rsCode.Update
rsService.AddNew
End Sub
Private Sub CmdDelete_Click()
rsService.Delete
Call ClearText(Me)
rsService.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsService.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsService.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsService.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsService.MoveNext
    If rsService.EOF = False Then
        Call RecordsetToText
    Else
        rsService.MoveFirst
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsService.MovePrevious
    If rsService.BOF = False Then
        Call RecordsetToText
    Else
        rsService.MoveLast
        Call RecordsetToText
    End If
End Sub
Private Sub CmdSave_Click()
Frame2.Enabled = True
Call TextToRecordset
rsService.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from ServiceMaster"
rsService.CursorLocation = adUseClient
rsService.Open q, cn, adOpenKeyset, adLockOptimistic
'**********
rsCode.CursorLocation = adUseClient
rsCode.Open "CodeSet", cn, adOpenKeyset, adLockOptimistic
End Sub
Public Function TextToRecordset()
rsService.Fields(0) = Val(TxtSCode.Text)
rsService.Fields(1) = IntoStr(Trim(TxtSName.Text))
rsService.Fields(2) = Val(TxtSprice.Text)
End Function
Public Function RecordsetToText()
 TxtSCode.Text = rsService.Fields(0)
 TxtSName.Text = rsService.Fields(1)
 TxtSprice.Text = rsService.Fields(2)
End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsService = Nothing
Set rsCode = Nothing
End Sub
Private Sub TxtSName_KeyPress(KeyAscii As Integer)
Call CheckName(KeyAscii)
End Sub
