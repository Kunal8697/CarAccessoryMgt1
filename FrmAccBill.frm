VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAccBill 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10200
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3525
      Left            =   1403
      TabIndex        =   21
      Top             =   990
      Width           =   7395
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2985
         Left            =   360
         TabIndex        =   22
         Top             =   180
         Width           =   6585
         Begin VB.TextBox TxtAmt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4140
            TabIndex        =   8
            Top             =   2250
            Width           =   1425
         End
         Begin VB.TextBox TxtDis 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1530
            TabIndex        =   7
            ToolTipText     =   "Enter Discount"
            Top             =   2250
            Width           =   1515
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4050
            TabIndex        =   32
            ToolTipText     =   "Select Bill Date"
            Top             =   450
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin VB.ComboBox CmbAccCode 
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
            Left            =   1500
            TabIndex        =   3
            ToolTipText     =   "Select Acc Code"
            Top             =   1350
            Width           =   1545
         End
         Begin VB.TextBox TxtCustName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3150
            TabIndex        =   2
            Top             =   900
            Width           =   3315
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
            Left            =   1500
            TabIndex        =   1
            ToolTipText     =   "Select Customer Code"
            Top             =   900
            Width           =   1545
         End
         Begin VB.TextBox TxtBillNo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1530
            TabIndex        =   0
            ToolTipText     =   "Enter Bill No"
            Top             =   450
            Width           =   1335
         End
         Begin VB.TextBox TxtAccName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3150
            TabIndex        =   4
            Top             =   1350
            Width           =   3315
         End
         Begin VB.TextBox TxtRate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1530
            TabIndex        =   5
            ToolTipText     =   "Enter Rate"
            Top             =   1800
            Width           =   1515
         End
         Begin VB.TextBox TxtQty 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4140
            TabIndex        =   6
            ToolTipText     =   "Enter Quantity"
            Top             =   1800
            Width           =   1425
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
            Left            =   3090
            TabIndex        =   31
            Top             =   540
            Width           =   900
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
            Left            =   3135
            TabIndex        =   30
            Top             =   2400
            Width           =   855
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
            Left            =   810
            TabIndex        =   29
            Top             =   1920
            Width           =   495
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
            Left            =   360
            TabIndex        =   28
            Top             =   1440
            Width           =   1005
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
            Left            =   195
            TabIndex        =   27
            Top             =   960
            Width           =   1110
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
            Left            =   630
            TabIndex        =   25
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Quantity"
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
            Left            =   3120
            TabIndex        =   24
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Discount"
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
            Left            =   360
            TabIndex        =   23
            Top             =   2340
            Width           =   945
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   1403
      TabIndex        =   9
      Top             =   4770
      Width           =   7395
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   180
         TabIndex        =   10
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   720
            Width           =   1275
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Accessories Bill Details"
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
      Left            =   3675
      TabIndex        =   26
      Top             =   360
      Width           =   2925
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2700
      Picture         =   "FrmAccBill.frx":0000
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "FrmAccBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim rsAccBill As New ADODB.Recordset
Dim rsCust As New ADODB.Recordset
Dim rsAcc As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset
Dim q As String
Dim dis As Integer
Dim Disval As Double
Private Sub CmbAccCode_Change()
If Not (CmbaccCode.Text = "") Then
        rsAcc.Filter = "Acode = " & CmbaccCode.Text
        If Not (rsAcc.EOF) Or Not (rsAcc.BOF) Then
            TxtAccName.Text = IntoStr(rsAcc.Fields(1))
                  End If
    End If
End Sub
Private Sub CmbAccCode_Click()
If Not (CmbaccCode.Text = "") Then
        rsAcc.Filter = "Acode = " & CmbaccCode.Text
        If Not (rsAcc.EOF) Or Not (rsAcc.BOF) Then
            TxtAccName.Text = IntoStr(rsAcc.Fields(1))
                   End If
    End If
End Sub
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
Private Sub CmdAdd_Click()
Call ClearText(Me)
Frame2.Enabled = True
Dim a As Integer
a = rsCode(3)
TxtBillNo.Text = a + 1
rsCode.Fields(3) = Val(TxtBillNo.Text)
rsCode.Update

rsAccBill.AddNew
End Sub
Private Sub CmdDelete_Click()
rsAccBill.Delete
Call ClearText(Me)
rsAccBill.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsAccBill.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsAccBill.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsAccBill.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsAccBill.MoveNext
    If rsAccBill.EOF = False Then
        Call RecordsetToText
    Else
        rsAccBill.MoveFirst
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsAccBill.MovePrevious
    If rsAccBill.BOF = False Then
        Call RecordsetToText
    Else
        rsAccBill.MoveLast
        Call RecordsetToText
    End If
End Sub



Private Sub CmdPrint_Click()
AccBill.Show
End Sub

Private Sub CmdSave_Click()
Frame2.Enabled = True
Call TextToRecordset
rsAccBill.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from ABillTran"
rsAccBill.CursorLocation = adUseClient
rsAccBill.Open q, cn, adOpenKeyset, adLockOptimistic
rsAcc.Open "AccessaryMaster", cn, adOpenKeyset, adLockOptimistic
rsCust.Open "CustTran", cn, adOpenKeyset, adLockOptimistic
'*************
rsCode.CursorLocation = adUseClient
rsCode.Open "CodeSet", cn, adOpenKeyset, adLockOptimistic
'***********
Call FillComboMenu
Call FillComboMenu1
End Sub
Private Sub FillComboMenu()
Dim j As Integer
       j = 0
    While rsAcc.EOF = False
        CmbaccCode.AddItem IntoStr(rsAcc.Fields(0))
        rsAcc.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbaccCode.ListIndex = 0
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
rsAccBill.Fields(0) = Val(TxtBillNo.Text)
rsAccBill.Fields(1) = IntoStr(Val(CmbCustCode.Text))
rsAccBill.Fields(2) = IntoStr(Val(CmbaccCode.Text))
rsAccBill.Fields(3) = IntoStr(Val(TxtQty.Text))
rsAccBill.Fields(4) = IntoStr(Val(TxtRate.Text))
rsAccBill.Fields(5) = IntoStr(Val(TxtDis.Text))
rsAccBill.Fields(6) = IntoStr(Val(TxtAmt.Text))
rsAccBill.Fields(7) = DTPicker1.Value
rsAccBill.Fields(8) = IntoStr(TxtAccName.Text)
End Function
Public Function RecordsetToText()
TxtBillNo.Text = rsAccBill.Fields(0)
CmbCustCode.Text = rsAccBill.Fields(1)
CmbaccCode.Text = rsAccBill.Fields(2)
TxtQty.Text = rsAccBill.Fields(3)
TxtRate.Text = rsAccBill.Fields(4)
TxtDis.Text = rsAccBill.Fields(5)
TxtAmt.Text = rsAccBill.Fields(6)
DTPicker1.Value = rsAccBill.Fields(7)
TxtAccName.Text = IntoStr(rsAccBill.Fields(8))
End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsAcc = Nothing
Set rsAccBill = Nothing
Set rsCust = Nothing
Set rsCode = Nothing

End Sub
Private Sub TxtDis_Click()
dis = InputBox("Enter Percentages of Discount")
TxtDis.Text = TxtRate.Text * TxtQty.Text * dis / 100
Disval = TxtRate.Text * TxtQty.Text - Val(TxtDis.Text)
TxtAmt.Text = Disval
TxtAmt.SetFocus
End Sub
