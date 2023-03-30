VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmServiceBill 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8550
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   578
      TabIndex        =   17
      Top             =   4320
      Width           =   7395
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   180
         TabIndex        =   18
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
            Height          =   375
            Left            =   5310
            Style           =   1  'Graphical
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   270
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   19
            Top             =   270
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   578
      TabIndex        =   9
      Top             =   840
      Width           =   7395
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2715
         Left            =   240
         TabIndex        =   10
         Top             =   180
         Width           =   6975
         Begin VB.TextBox TxtRate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   6
            ToolTipText     =   "Enter Rate"
            Top             =   1800
            Width           =   1515
         End
         Begin VB.TextBox TxtSName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1350
            Width           =   3315
         End
         Begin VB.TextBox TxtBillNo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   0
            ToolTipText     =   "Enter Bill No"
            Top             =   450
            Width           =   1515
         End
         Begin VB.ComboBox CmbCustCode 
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
            Left            =   1590
            TabIndex        =   2
            ToolTipText     =   "Select Cust Code"
            Top             =   900
            Width           =   1545
         End
         Begin VB.TextBox TxtCustName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   900
            Width           =   3315
         End
         Begin VB.ComboBox CmbSerCode 
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
            Left            =   1590
            TabIndex        =   4
            ToolTipText     =   "Select Service Code"
            Top             =   1350
            Width           =   1545
         End
         Begin VB.TextBox TxtAmt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4410
            TabIndex        =   7
            ToolTipText     =   "Enter Amount"
            Top             =   1890
            Width           =   1425
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4320
            TabIndex        =   1
            ToolTipText     =   "Enter Bill Date"
            Top             =   450
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Bill No"
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
            Left            =   840
            TabIndex        =   16
            Top             =   450
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Cust Code"
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
            Left            =   405
            TabIndex        =   15
            Top             =   900
            Width           =   1110
         End
         Begin VB.Label Label6 
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
            Left            =   90
            TabIndex        =   14
            Top             =   1350
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Rate"
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
            Left            =   1020
            TabIndex        =   13
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Amount"
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
            Left            =   3510
            TabIndex        =   12
            Top             =   1890
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Bill Date"
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
            Left            =   3240
            TabIndex        =   11
            Top             =   540
            Width           =   900
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2625
      Picture         =   "FrmServiceBill.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Service Bill Details"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   270
      Width           =   2355
   End
End
Attribute VB_Name = "FrmServiceBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsServiceBill As New ADODB.Recordset
Dim rsCust As New ADODB.Recordset
Dim rsSer As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset
Dim q As String
Private Sub CmbCustCode_Change()
If Not (CmbCustCode.Text = "") Then
        rsCust.Filter = "CustCode = " & CmbCustCode.Text
        If Not (rsCust.EOF) Or Not (rsCust.BOF) Then
            TxtCustName.Text = IntoStr(rsCust.Fields(1))
        End If
    End If
End Sub
Private Sub CmbCustCode_Click()
If Not (CmbCustCode.Text = "") Then
        rsCust.Filter = "CustCode = " & CmbCustCode.Text
        If Not (rsCust.EOF) Or Not (rsCust.BOF) Then
            TxtCustName.Text = IntoStr(rsCust.Fields(1))
        End If
    End If
End Sub
Private Sub CmbSerCode_Change()
If Not (CmbSerCode.Text = "") Then
        rsSer.Filter = "Scode = " & CmbSerCode.Text
        If Not (rsSer.EOF) Or Not (rsSer.BOF) Then
            TxtSName.Text = IntoStr(rsSer.Fields(1))
            TxtRate.Text = IntoStr(rsSer.Fields(2))
        End If
    End If
End Sub
Private Sub CmbSerCode_Click()
If Not (CmbSerCode.Text = "") Then
        rsSer.Filter = "Scode = " & CmbSerCode.Text
        If Not (rsSer.EOF) Or Not (rsSer.BOF) Then
            TxtSName.Text = IntoStr(rsSer.Fields(1))
            TxtRate.Text = IntoStr(rsSer.Fields(2))
        End If
    End If
End Sub
Private Sub CmdAdd_Click()
Frame2.Enabled = True
Call ClearText(Me)
Dim a As Integer
a = rsCode(4)
TxtBillNo.Text = a + 1
rsCode.Fields(4) = Val(TxtBillNo.Text)
rsCode.Update

rsServiceBill.AddNew
End Sub
Private Sub CmdDelete_Click()
rsServiceBill.Delete
Call ClearText(Me)
rsServiceBill.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsServiceBill.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsServiceBill.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsServiceBill.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsServiceBill.MoveNext
    If rsServiceBill.EOF = False Then
        Call RecordsetToText
    Else
        rsServiceBill.MoveFirst
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsServiceBill.MovePrevious
    If rsServiceBill.BOF = False Then
        Call RecordsetToText
    Else
        rsServiceBill.MoveLast
        Call RecordsetToText
    End If
End Sub



Private Sub CmdPrint_Click()
SerBill.Show
End Sub

Private Sub CmdSave_Click()
Frame2.Enabled = True
Call TextToRecordset
rsServiceBill.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from SBillTran"
rsServiceBill.CursorLocation = adUseClient
rsServiceBill.Open q, cn, adOpenKeyset, adLockOptimistic
rsSer.Open "ServiceMaster", cn, adOpenKeyset, adLockOptimistic
rsCust.Open "CustTran", cn, adOpenKeyset, adLockOptimistic
'**********
rsCode.CursorLocation = adUseClient
rsCode.Open "CodeSet", cn, adOpenKeyset, adLockOptimistic
Call FillComboMenu
Call FillComboMenu1
 End Sub
 Private Sub FillComboMenu()
Dim j As Integer
       j = 0
    While rsSer.EOF = False
        CmbSerCode.AddItem IntoStr(rsSer.Fields(0))
        rsSer.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbSerCode.ListIndex = 0
    End If
End Sub
Private Sub FillComboMenu1()

Dim j As Integer
       j = 0
    While rsCust.EOF = False
        CmbCustCode.AddItem IntoStr(rsCust.Fields(0))
        rsCust.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbCustCode.ListIndex = 0
    End If
End Sub
Public Function TextToRecordset()
rsServiceBill.Fields(0) = Val(TxtBillNo.Text)
rsServiceBill.Fields(1) = IntoStr(Val(CmbCustCode.Text))
rsServiceBill.Fields(2) = IntoStr(Val(CmbSerCode.Text))
rsServiceBill.Fields(3) = IntoStr(Val(TxtRate.Text))
rsServiceBill.Fields(4) = IntoStr(Val(TxtAmt.Text))
rsServiceBill.Fields(5) = DTPicker1.Value
End Function
Public Function RecordsetToText()
TxtBillNo.Text = rsServiceBill.Fields(0)
CmbCustCode.Text = rsServiceBill.Fields(1)
CmbSerCode.Text = rsServiceBill.Fields(2)
TxtRate.Text = rsServiceBill.Fields(3)
TxtAmt.Text = rsServiceBill.Fields(4)
DTPicker1.Value = rsServiceBill.Fields(5)
End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsCust = Nothing
Set rsSer = Nothing
Set rsServiceBill = Nothing
Set rsCode = Nothing
End Sub
