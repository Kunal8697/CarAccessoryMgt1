VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmQuatation 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8625
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   570
      TabIndex        =   22
      Top             =   900
      Width           =   7395
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   3885
         Left            =   360
         TabIndex        =   23
         Top             =   180
         Width           =   6855
         Begin VB.TextBox TxtAccName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3480
            TabIndex        =   8
            Top             =   2880
            Width           =   2595
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Left            =   1530
            TabIndex        =   5
            ToolTipText     =   "Enter Phone No"
            Top             =   2340
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "##-####-#######"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox CmbCity 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "FrmQuatation.frx":0000
            Left            =   1530
            List            =   "FrmQuatation.frx":0013
            TabIndex        =   4
            ToolTipText     =   "Select City"
            Top             =   1800
            Width           =   1545
         End
         Begin VB.TextBox TxtAmt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1530
            TabIndex        =   9
            ToolTipText     =   "Enter Total Cost"
            Top             =   3330
            Width           =   1875
         End
         Begin VB.ComboBox CmbaccCode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1530
            TabIndex        =   7
            ToolTipText     =   "Select Accessory Code"
            Top             =   2880
            Width           =   1905
         End
         Begin VB.TextBox TxtCustName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1530
            TabIndex        =   2
            ToolTipText     =   "Enter Name Of Customer"
            Top             =   900
            Width           =   3315
         End
         Begin VB.TextBox TxtSrNo 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1530
            TabIndex        =   0
            ToolTipText     =   "Enter Serial No"
            Top             =   450
            Width           =   1515
         End
         Begin VB.TextBox TxtAdd 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1530
            TabIndex        =   3
            ToolTipText     =   "Enter Address"
            Top             =   1350
            Width           =   3315
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4320
            TabIndex        =   1
            ToolTipText     =   "Select Current Date"
            Top             =   450
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   375
            Left            =   3510
            TabIndex        =   6
            ToolTipText     =   "Enter Phone No"
            Top             =   2340
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "############"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Acc Code"
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
            Index           =   3
            Left            =   210
            TabIndex        =   32
            Top             =   2880
            Width           =   1005
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Ph.No"
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
            Index           =   2
            Left            =   585
            TabIndex        =   31
            Top             =   2340
            Width           =   630
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "City"
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
            Index           =   1
            Left            =   780
            TabIndex        =   30
            Top             =   1830
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Sr.No"
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
            Left            =   630
            TabIndex        =   29
            Top             =   450
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   " Date"
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
            Left            =   3690
            TabIndex        =   27
            Top             =   540
            Width           =   570
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Total"
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
            Left            =   675
            TabIndex        =   26
            Top             =   3330
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Address"
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
            Index           =   0
            Left            =   330
            TabIndex        =   25
            Top             =   1350
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name"
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
            Left            =   585
            TabIndex        =   24
            Top             =   900
            Width           =   630
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   660
      TabIndex        =   10
      Top             =   5400
      Width           =   7395
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   180
         TabIndex        =   11
         Top             =   90
         Width           =   7035
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton CmdPrint 
            BackColor       =   &H00C0E0FF&
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
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   270
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
            Height          =   375
            Left            =   5310
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   720
            Width           =   1275
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Quatation Of Accessories"
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
      Left            =   2760
      TabIndex        =   28
      Top             =   360
      Width           =   3300
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2070
      Picture         =   "FrmQuatation.frx":0040
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "FrmQuatation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsQuat As New ADODB.Recordset
Dim rsAMast As New ADODB.Recordset
Dim q As String
Private Sub CmbAccCode_Change()
If Not (CmbaccCode.Text = "") Then
        rsAMast.Filter = "Acode = " & CmbaccCode.Text
        If Not (rsAMast.EOF) Or Not (rsAMast.BOF) Then
            TxtAccName.Text = IntoStr(rsAMast.Fields(1))
            TxtAmt.Text = IntoStr(rsAMast.Fields(4))
         End If
    End If
End Sub
Private Sub CmbAccCode_Click()
If Not (CmbaccCode.Text = "") Then
        rsAMast.Filter = "Acode = " & CmbaccCode.Text
        If Not (rsAMast.EOF) Or Not (rsAMast.BOF) Then
            TxtAccName.Text = IntoStr(rsAMast.Fields(1))
           TxtAmt.Text = IntoStr(rsAMast.Fields(4))
         End If
    End If
End Sub
Private Sub CmdAdd_Click()
Frame2.Enabled = True
Call ClearText(Me)
rsQuat.AddNew
End Sub
Private Sub CmdDelete_Click()
rsQuat.Delete
Call ClearText(Me)
rsQuat.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsQuat.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsQuat.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsQuat.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsQuat.MoveNext
    If rsQuat.EOF = False Then
        Call RecordsetToText
    Else
        rsQuat.MoveFirst
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsQuat.MovePrevious
    If rsQuat.BOF = False Then
        Call RecordsetToText
    Else
        rsQuat.MoveLast
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrint_Click()
Quat.Show
End Sub
Private Sub CmdSave_Click()
Frame2.Enabled = True
Call TextToRecordset
rsQuat.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from Quotation"
rsQuat.CursorLocation = adUseClient
rsQuat.Open q, cn, adOpenKeyset, adLockOptimistic
rsAMast.Open "AccessaryMaster", cn, adOpenKeyset, adLockOptimistic
Call FillComboMenu
End Sub
Private Sub FillComboMenu()
Dim j As Integer
      j = 0
    While rsAMast.EOF = False
        CmbaccCode.AddItem IntoStr(rsAMast.Fields(0))
        rsAMast.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbaccCode.ListIndex = 0
    End If
     
End Sub
Public Function TextToRecordset()
rsQuat.Fields(0) = Val(TxtSrNo.Text)
rsQuat.Fields(1) = IntoStr(TxtCustName.Text)
rsQuat.Fields(2) = IntoStr(TxtAdd.Text)
rsQuat.Fields(3) = IntoStr(CmbCity.Text)
rsQuat.Fields(4) = DTPicker1.Value
rsQuat.Fields(5) = MaskEdBox1.Text
rsQuat.Fields(6) = MaskEdBox2.Text
rsQuat.Fields(7) = Val(CmbaccCode.Text)
rsQuat.Fields(8) = IntoStr(TxtAccName.Text)
rsQuat.Fields(9) = IntoStr(TxtAmt.Text)
End Function
Public Function RecordsetToText()
TxtSrNo.Text = rsQuat.Fields(0)
TxtCustName.Text = rsQuat.Fields(1)
TxtAdd.Text = rsQuat.Fields(2)
CmbCity.Text = rsQuat.Fields(3)
DTPicker1.Value = rsQuat.Fields(4)
MaskEdBox1.Text = rsQuat.Fields(5)
MaskEdBox2.Text = rsQuat.Fields(6)
CmbaccCode.Text = rsQuat.Fields(7)
TxtAccName.Text = rsQuat.Fields(8)
TxtAmt.Text = rsQuat.Fields(9)
End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsAMast = Nothing
Set rsQuat = Nothing
End Sub
Private Sub TxtCustName_KeyPress(KeyAscii As Integer)
Call CheckName(KeyAscii)
End Sub
