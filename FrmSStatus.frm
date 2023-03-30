VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSStatus 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7815
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   270
      TabIndex        =   16
      Top             =   5400
      Width           =   7335
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   180
         TabIndex        =   17
         Top             =   90
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
            Left            =   3945
            Style           =   1  'Graphical
            TabIndex        =   26
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
            Left            =   2685
            Style           =   1  'Graphical
            TabIndex        =   25
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
            Left            =   165
            Style           =   1  'Graphical
            TabIndex        =   24
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
            Left            =   1425
            Style           =   1  'Graphical
            TabIndex        =   23
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
            Left            =   165
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
            Left            =   1425
            Style           =   1  'Graphical
            TabIndex        =   21
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
            Left            =   3945
            Style           =   1  'Graphical
            TabIndex        =   20
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
            Left            =   2685
            Style           =   1  'Graphical
            TabIndex        =   19
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
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4515
      Left            =   278
      TabIndex        =   1
      Top             =   720
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   4155
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   6945
         Begin VB.TextBox TxtSName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   4665
            TabIndex        =   29
            Top             =   1890
            Width           =   2055
         End
         Begin VB.ComboBox CmbSCode 
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
            Left            =   2415
            TabIndex        =   28
            ToolTipText     =   "Select Service Code"
            Top             =   1890
            Width           =   1995
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
            Left            =   2430
            TabIndex        =   27
            ToolTipText     =   "Select Cust Code"
            Top             =   360
            Width           =   1545
         End
         Begin VB.TextBox TxtAccName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   4680
            TabIndex        =   9
            Top             =   1350
            Width           =   1965
         End
         Begin VB.TextBox TxtCustName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2430
            TabIndex        =   8
            Top             =   810
            Width           =   4215
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
            Left            =   2430
            TabIndex        =   7
            ToolTipText     =   "Select Accessory Code"
            Top             =   1350
            Width           =   1995
         End
         Begin VB.TextBox TxtNoOfService 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2430
            TabIndex        =   6
            ToolTipText     =   "Enter No Of Services"
            Top             =   2430
            Width           =   2055
         End
         Begin VB.TextBox TxTAmt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2430
            TabIndex        =   5
            ToolTipText     =   "Enter Amount"
            Top             =   2970
            Width           =   2055
         End
         Begin VB.TextBox TxtNAmt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2430
            TabIndex        =   4
            ToolTipText     =   "Enter Net Amount"
            Top             =   3510
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4950
            TabIndex        =   3
            ToolTipText     =   "Enter Current Date"
            Top             =   270
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin VB.Label Label5 
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
            Left            =   1470
            TabIndex        =   31
            Top             =   3060
            Width           =   855
         End
         Begin VB.Label Label4 
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
            Left            =   915
            TabIndex        =   30
            Top             =   1890
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name Of Customer"
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
            Left            =   300
            TabIndex        =   15
            Top             =   900
            Width           =   2055
         End
         Begin VB.Label Label2 
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
            Left            =   1200
            TabIndex        =   14
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label8 
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
            Left            =   465
            TabIndex        =   13
            Top             =   1350
            Width           =   1920
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "No Of Services"
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
            Left            =   720
            TabIndex        =   12
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label12 
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
            Left            =   1035
            TabIndex        =   11
            Top             =   3510
            Width           =   1290
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Date"
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
            Left            =   4230
            TabIndex        =   10
            Top             =   270
            Width           =   510
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2100
      Picture         =   "FrmSStatus.frx":0000
      Top             =   180
      Width           =   480
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
      Left            =   2910
      TabIndex        =   0
      Top             =   270
      Width           =   1920
   End
End
Attribute VB_Name = "FrmSStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim q As String
Dim rsSTran As New ADODB.Recordset
Dim rsAMast As New ADODB.Recordset
Dim rsService As New ADODB.Recordset
Dim rsCust As New ADODB.Recordset
Private Sub CmbAccCode_Change()
If Not (CmbaccCode.Text = "") Then
        rsAMast.Filter = "Acode = " & CmbaccCode.Text
        If Not (rsAMast.EOF) Or Not (rsAMast.BOF) Then
            TxtAccName.Text = IntoStr(rsAMast.Fields(1))
            
        End If
    End If

End Sub

Private Sub CmbAccCode_Click()
If Not (CmbaccCode.Text = "") Then
        rsAMast.Filter = "Acode = " & CmbaccCode.Text
        If Not (rsAMast.EOF) Or Not (rsAMast.BOF) Then
            TxtAccName.Text = IntoStr(rsAMast.Fields(1))
            
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

Private Sub CmbSCode_Change()
If Not (CmbSCode.Text = "") Then
        rsService.Filter = "Scode = " & CmbSCode.Text
        If Not (rsService.EOF) Or Not (rsService.BOF) Then
            TxtSName.Text = IntoStr(rsService.Fields(1))
        End If
    End If
End Sub

Private Sub CmbSCode_Click()
If Not (CmbSCode.Text = "") Then
        rsService.Filter = "Scode = " & CmbSCode.Text
        If Not (rsService.EOF) Or Not (rsService.BOF) Then
            TxtSName.Text = IntoStr(rsService.Fields(1))
        End If
    End If
End Sub

Private Sub CmdAdd_Click()
Frame2.Enabled = True
Call ClearText(Me)
rsSTran.AddNew

End Sub

Private Sub CmdDelete_Click()
rsSTran.Delete
Call ClearText(Me)
rsSTran.MoveFirst
Call RecordsetToText
End Sub

Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsSTran.Update

End Sub

Private Sub CmdExit_Click()
Unload Me

End Sub

Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsSTran.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsSTran.MoveLast
Call RecordsetToText

End Sub

Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsSTran.MoveNext
    If rsSTran.EOF = False Then
        Call RecordsetToText
    Else
        rsSTran.MoveFirst
        Call RecordsetToText
    End If

End Sub

Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsSTran.MovePrevious

    If rsSTran.BOF = False Then
        Call RecordsetToText
    Else
        rsSTran.MoveLast
        Call RecordsetToText
    End If

End Sub




Private Sub CmdSave_Click()
Frame2.Enabled = True
Call TextToRecordset
rsSTran.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub

Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from SStatusTran"
rsSTran.CursorLocation = adUseClient
rsSTran.Open q, cn, adOpenKeyset, adLockOptimistic
rsAMast.Open "AccessaryMaster", cn, adOpenKeyset, adLockOptimistic
rsService.Open "ServiceMaster", cn, adOpenKeyset, adLockOptimistic
rsCust.Open "CustTran", cn, adOpenKeyset, adLockOptimistic
Call FillComboMenu
Call FillComboMenu1
Call FillComboMenu2

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

Private Sub FillComboMenu1()
Dim j As Integer
       j = 0
    While rsService.EOF = False
        CmbSCode.AddItem IntoStr(rsService.Fields(0))
        rsService.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbSCode.ListIndex = 0
    End If
End Sub
Public Function TextToRecordset()
rsSTran.Fields(0) = Val(CmbCustCode.Text)
rsSTran.Fields(1) = Val(CmbaccCode.Text)
rsSTran.Fields(2) = Val(CmbSCode.Text)
rsSTran.Fields(3) = Val(TxtNoOfService.Text)
rsSTran.Fields(4) = Val(TxtAmt.Text)
rsSTran.Fields(5) = Val(TxtNAmt.Text)
rsSTran.Fields(6) = DTPicker1.Value
End Function
Public Function RecordsetToText()
 CmbCustCode.Text = rsSTran.Fields(0)
 CmbaccCode.Text = rsSTran.Fields(1)
 CmbSCode.Text = rsSTran.Fields(2)
 TxtNoOfService.Text = rsSTran.Fields(3)
 TxtAmt.Text = rsSTran.Fields(4)
 TxtNAmt.Text = rsSTran.Fields(5)
 DTPicker1.Value = rsSTran.Fields(6)
 
End Function

Private Sub FillComboMenu2()
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

Private Sub Form_Unload(Cancel As Integer)
Set rsAMast = Nothing
Set rsCust = Nothing
Set rsService = Nothing
Set rsSTran = Nothing

End Sub

Private Sub TxTAmt_LostFocus()
TxtNAmt.Text = TxtNoOfService.Text * TxtAmt.Text

End Sub

Private Sub TxtCustName_KeyPress(KeyAscii As Integer)
Call CheckName(KeyAscii)
End Sub
