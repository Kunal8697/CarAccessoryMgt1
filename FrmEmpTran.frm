VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEmpTran 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8610
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   608
      TabIndex        =   14
      Top             =   4500
      Width           =   7395
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   270
         TabIndex        =   15
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   16
            Top             =   270
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3435
      Left            =   653
      TabIndex        =   9
      Top             =   900
      Width           =   7305
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2895
         Left            =   270
         TabIndex        =   10
         Top             =   360
         Width           =   6495
         Begin VB.TextBox TxtNsal 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1620
            Width           =   1335
         End
         Begin VB.TextBox TxtEarnings 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox TxtDed 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4590
            TabIndex        =   5
            Top             =   1260
            Width           =   1335
         End
         Begin VB.TextBox TxtAb 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4590
            TabIndex        =   3
            ToolTipText     =   "Enter Days Of Absent"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtBpay 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox CmbEmpCode 
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
            TabIndex        =   0
            ToolTipText     =   "Select Employee Code"
            Top             =   270
            Width           =   1455
         End
         Begin VB.TextBox TxtEName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3060
            TabIndex        =   1
            Top             =   270
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker DTPPayDate 
            Height          =   375
            Left            =   2250
            TabIndex        =   7
            ToolTipText     =   "Select Date Of Payments"
            Top             =   2160
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Net Sal"
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
            Left            =   600
            TabIndex        =   29
            Top             =   1710
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Date  Of Payments"
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
            Left            =   150
            TabIndex        =   28
            Top             =   2160
            Width           =   2025
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Deductions"
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
            TabIndex        =   27
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Earnings"
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
            Left            =   420
            TabIndex        =   26
            Top             =   1260
            Width           =   945
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
            Left            =   270
            TabIndex        =   13
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Absent Day's"
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
            Left            =   3030
            TabIndex        =   12
            Top             =   720
            Width           =   1425
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
            Left            =   315
            TabIndex        =   11
            Top             =   720
            Width           =   1050
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2700
      Picture         =   "FrmEmpTran.frx":0000
      Top             =   270
      Width           =   480
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
      Left            =   3240
      TabIndex        =   8
      Top             =   360
      Width           =   2235
   End
End
Attribute VB_Name = "FrmEmpTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmp As New ADODB.Recordset
Dim rsEmpTran As New ADODB.Recordset
Dim q As String
Dim ab As Integer
Private Sub CmbEmpCode_Change()
If Not (CmbEmpCode.Text = "") Then
        rsEmp.Filter = "EmpCode = " & CmbEmpCode.Text
        If Not (rsEmp.EOF) Or Not (rsEmp.BOF) Then
              TxtEName.Text = IntoStr(rsEmp.Fields(1))
        TxtBpay.Text = IntoStr(rsEmp.Fields(3))
        End If
    End If
End Sub

Private Sub CmbEmpCode_Click()
If Not (CmbEmpCode.Text = "") Then
        rsEmp.Filter = "EmpCode = " & CmbEmpCode.Text
        If Not (rsEmp.EOF) Or Not (rsEmp.BOF) Then
              TxtEName.Text = IntoStr(rsEmp.Fields(1))
        TxtBpay.Text = IntoStr(rsEmp.Fields(3))
        End If
    End If
End Sub

Private Sub CmdAdd_Click()
Call ClearText(Me)
Frame2.Enabled = True
rsEmpTran.AddNew
End Sub

Private Sub CmdDelete_Click()
rsEmpTran.Delete
Call ClearText(Me)
rsEmpTran.MoveFirst
Call RecordsetToText
End Sub

Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsEmpTran.Update

End Sub

Private Sub CmdExit_Click()
Unload Me

End Sub

Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsEmpTran.MoveFirst
Call RecordsetToText
End Sub

Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsEmpTran.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsEmpTran.MoveNext
    If rsEmpTran.EOF = False Then
        Call RecordsetToText
    Else
        rsEmpTran.MoveFirst
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsEmpTran.MovePrevious

    If rsEmpTran.BOF = False Then
        Call RecordsetToText
    Else
        rsEmpTran.MoveLast
        Call RecordsetToText
    End If

End Sub




Private Sub CmdPrint_Click()
CustTrans.Show
End Sub

Private Sub CmdSave_Click()
Frame2.Enabled = False
Call TextToRecordset
rsEmpTran.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub

Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from EmpTran"
rsEmpTran.CursorLocation = adUseClient
rsEmpTran.Open q, cn, adOpenKeyset, adLockOptimistic
rsEmp.Open "EmpMaster", cn, adOpenKeyset, adLockOptimistic
Call FillComboMenu
End Sub
Private Sub FillComboMenu()
Dim j As Integer
       j = 0
    While rsEmp.EOF = False
        CmbEmpCode.AddItem IntoStr(rsEmp.Fields(0))
        rsEmp.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbEmpCode.ListIndex = 0
    End If
End Sub
Public Function TextToRecordset()
rsEmpTran.Fields(0) = Val(CmbEmpCode.Text)
rsEmpTran.Fields(1) = Val(TxtBpay.Text)
rsEmpTran.Fields(2) = Val(TxtAb.Text)
rsEmpTran.Fields(3) = Val(TxtEarnings.Text)
rsEmpTran.Fields(4) = Val(TxtDed.Text)
rsEmpTran.Fields(5) = Val(TxtNsal.Text)
rsEmpTran.Fields(6) = DTPPayDate.Value
End Function
Public Function RecordsetToText()
CmbEmpCode.Text = rsEmpTran.Fields(0)
TxtBpay.Text = rsEmpTran.Fields(1)
TxtAb.Text = rsEmpTran.Fields(2)
TxtEarnings.Text = rsEmpTran.Fields(3)
TxtDed.Text = rsEmpTran.Fields(4)
TxtNsal.Text = rsEmpTran.Fields(5)
DTPPayDate.Value = rsEmpTran.Fields(6)
End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsEmp = Nothing
Set rsEmpTran = Nothing
End Sub
Private Sub TxtAb_LostFocus()
ab = TxtBpay.Text / 30
TxtDed.Text = TxtAb.Text * ab
TxtNsal.Text = Val(TxtEarnings.Text) - Val(TxtDed.Text)
End Sub
Private Sub TxtBpay_Change()
TxtEarnings.Text = TxtBpay.Text
End Sub
Private Sub TxtEName_KeyPress(KeyAscii As Integer)
Call CheckName(KeyAscii)
End Sub
