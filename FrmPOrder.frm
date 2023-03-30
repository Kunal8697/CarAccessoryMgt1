VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPOrder 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   735
      TabIndex        =   10
      Top             =   5220
      Width           =   7665
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   450
         TabIndex        =   11
         Top             =   90
         Width           =   6765
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   12
            Top             =   270
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4185
      Left            =   735
      TabIndex        =   0
      Top             =   720
      Width           =   7665
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   3765
         Left            =   360
         TabIndex        =   1
         Top             =   180
         Width           =   6945
         Begin MSComCtl2.DTPicker DTPCDate 
            Height          =   375
            Left            =   5220
            TabIndex        =   31
            ToolTipText     =   "Enter Current Date"
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin VB.TextBox TxtCdays 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1710
            TabIndex        =   30
            ToolTipText     =   "Enter Credit Day's"
            Top             =   3060
            Width           =   1695
         End
         Begin VB.TextBox TxtAmt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4950
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   2430
            Width           =   1605
         End
         Begin VB.TextBox TxtDis 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1710
            TabIndex        =   28
            Top             =   2520
            Width           =   1665
         End
         Begin VB.TextBox TxtAName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3510
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1350
            Width           =   2955
         End
         Begin VB.ComboBox CmbAcode 
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
            Left            =   1710
            TabIndex        =   22
            ToolTipText     =   "Select Acc Code"
            Top             =   1350
            Width           =   1545
         End
         Begin VB.TextBox TxtRate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   4950
            Locked          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Enter Rate"
            Top             =   1890
            Width           =   1605
         End
         Begin VB.TextBox TxtPno 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1710
            TabIndex        =   4
            ToolTipText     =   "Enter P Order No"
            Top             =   270
            Width           =   1515
         End
         Begin VB.TextBox TxtQty 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1710
            TabIndex        =   3
            ToolTipText     =   "Enter Quantity"
            Top             =   1980
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker DTPODate 
            Height          =   375
            Left            =   1710
            TabIndex        =   2
            Top             =   810
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Current Date"
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
            Left            =   3570
            TabIndex        =   32
            Top             =   270
            Width           =   1395
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Credit Day's"
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
            Left            =   120
            TabIndex        =   27
            Top             =   3150
            Width           =   1320
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Net Amount"
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
            Left            =   3555
            TabIndex        =   26
            Top             =   2430
            Width           =   1290
         End
         Begin VB.Label Label6 
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
            Left            =   495
            TabIndex        =   25
            Top             =   2610
            Width           =   945
         End
         Begin VB.Label Label1 
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
            Left            =   4350
            TabIndex        =   24
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Order Date"
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
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "P.Order No"
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
            Left            =   270
            TabIndex        =   8
            Top             =   270
            Width           =   1170
         End
         Begin VB.Label Label4 
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
            Left            =   435
            TabIndex        =   7
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label5 
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
            Left            =   495
            TabIndex        =   6
            Top             =   1980
            Width           =   945
         End
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Purchase Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   3105
      TabIndex        =   33
      Top             =   180
      Width           =   2190
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2340
      Picture         =   "FrmPOrder.frx":0000
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "FrmPOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPorder As New ADODB.Recordset
Dim rsAcc As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset
Dim q As String
Dim a As Double

Private Sub CmbAcode_Change()
If Not (CmbAcode.Text = "") Then
        rsAcc.Filter = "Acode = " & CmbAcode.Text
        If Not (rsAcc.EOF) Or Not (rsAcc.BOF) Then
            TxtAName.Text = IntoStr(rsAcc.Fields(1))
            TxtRate.Text = IntoStr(rsAcc.Fields(4))
            
        End If
    End If
End Sub

Private Sub CmbAcode_Click()

If Not (CmbAcode.Text = "") Then
        rsAcc.Filter = "Acode = " & CmbAcode.Text
        If Not (rsAcc.EOF) Or Not (rsAcc.BOF) Then
            TxtAName.Text = IntoStr(rsAcc.Fields(1))
            TxtRate.Text = IntoStr(rsAcc.Fields(4))
        End If
    End If
End Sub

Private Sub CmdAdd_Click()
Frame2.Enabled = True
Call ClearText(Me)
Dim a As Integer
a = rsCode(5)
TxtPno.Text = a + 1
rsCode.Fields(5) = Val(TxtPno.Text)
rsCode.Update
rsPorder.AddNew
End Sub

Private Sub CmdDelete_Click()
rsPorder.Delete
Call ClearText(Me)
rsPorder.MoveFirst
Call RecordsetToText
End Sub

Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsPorder.Update

End Sub

Private Sub CmdExit_Click()
Unload Me

End Sub

Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsPorder.MoveFirst
Call RecordsetToText

End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsPorder.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsPorder.MoveNext
    If rsPorder.EOF = False Then
        Call RecordsetToText
    Else
        rsPorder.MoveFirst
        Call RecordsetToText
    End If

End Sub

Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsPorder.MovePrevious

    If rsPorder.BOF = False Then
        Call RecordsetToText
    Else
        rsPorder.MoveLast
        Call RecordsetToText
    End If

End Sub




Private Sub CmdPrint_Click()
PurTran.Show
End Sub
Private Sub CmdSave_Click()
Frame2.Enabled = False
Call TextToRecordset
rsPorder.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from PurchaseTran"
rsPorder.CursorLocation = adUseClient
rsPorder.Open q, cn, adOpenKeyset, adLockOptimistic
rsAcc.Open "AccessaryMaster", cn, adOpenKeyset, adLockBatchOptimistic

'************
rsCode.CursorLocation = adUseClient
rsCode.Open "CodeSet", cn, adOpenKeyset, adLockOptimistic
Call FillComboMenu
End Sub
Public Function TextToRecordset()
rsPorder.Fields(0) = Val(TxtPno.Text)
rsPorder.Fields(1) = DTPODate.Value
rsPorder.Fields(2) = IntoStr(Val(CmbAcode.Text))
rsPorder.Fields(3) = IntoStr(Val(TxtQty.Text))
rsPorder.Fields(4) = IntoStr(Val(TxtRate.Text))
rsPorder.Fields(5) = IntoStr(Val(TxtDis.Text))
rsPorder.Fields(6) = IntoStr(Val(TxtAmt.Text))
rsPorder.Fields(7) = IntoStr(Val(TxtCdays.Text))
rsPorder.Fields(8) = DTPCDate.Value
End Function

Public Function RecordsetToText()
  TxtPno.Text = rsPorder.Fields(0)
DTPODate.Value = rsPorder.Fields(1)
CmbAcode.Text = rsPorder.Fields(2)
TxtQty.Text = rsPorder.Fields(3)
TxtRate.Text = rsPorder.Fields(4)
 TxtDis.Text = rsPorder.Fields(5)
 TxtAmt.Text = rsPorder.Fields(6)
 TxtCdays.Text = rsPorder.Fields(7)
 DTPCDate.Value = rsPorder.Fields(8)
End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsPorder = Nothing
Set rsAcc = Nothing
Set rsCode = Nothing
End Sub

Private Sub TxtDis_LostFocus()
a = TxtQty.Text * TxtRate.Text
MsgBox (a)
TxtDis.Text = a * 5 / 100
TxtAmt.Text = a - TxtDis.Text
End Sub

Private Sub FillComboMenu()
Dim j As Integer
 
    j = 0
    While rsAcc.EOF = False
        CmbAcode.AddItem IntoStr(rsAcc.Fields(0))
        rsAcc.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbAcode.ListIndex = 0
    End If
    
    
End Sub
