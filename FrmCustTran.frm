VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCustTran 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10065
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5685
      Left            =   525
      TabIndex        =   18
      Top             =   810
      Width           =   9015
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   5415
         Left            =   90
         TabIndex        =   19
         Top             =   120
         Width           =   8745
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4950
            TabIndex        =   1
            ToolTipText     =   "Enter Current Date"
            Top             =   270
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   43355
         End
         Begin VB.ComboBox CmbService 
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
            ItemData        =   "FrmCustTran.frx":0000
            Left            =   2520
            List            =   "FrmCustTran.frx":000A
            TabIndex        =   15
            ToolTipText     =   "Select Service Status"
            Top             =   4860
            Width           =   1995
         End
         Begin VB.TextBox TxtNAmt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   4500
            Width           =   1785
         End
         Begin VB.TextBox TxtCdays 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2520
            TabIndex        =   13
            ToolTipText     =   "Enter Credit Day's"
            Top             =   4410
            Width           =   2055
         End
         Begin VB.TextBox TxtDis 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   6120
            TabIndex        =   12
            Top             =   3960
            Width           =   1785
         End
         Begin VB.TextBox TxtQty 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2520
            TabIndex        =   11
            ToolTipText     =   "Enter Quantity Of Accessories"
            Top             =   3870
            Width           =   2055
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
            Left            =   2520
            TabIndex        =   9
            ToolTipText     =   "Select Acc Code"
            Top             =   3330
            Width           =   1995
         End
         Begin MSMask.MaskEdBox MaskPhone 
            Height          =   375
            Left            =   2520
            TabIndex        =   7
            ToolTipText     =   "Enter Phone No"
            Top             =   2790
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "####-#######"
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
            ItemData        =   "FrmCustTran.frx":0017
            Left            =   2520
            List            =   "FrmCustTran.frx":002D
            TabIndex        =   6
            ToolTipText     =   "Select N ame Of City"
            Top             =   2340
            Width           =   1995
         End
         Begin VB.TextBox TxtAdd1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2520
            TabIndex        =   5
            ToolTipText     =   "Enter Additional Address"
            Top             =   1800
            Width           =   4215
         End
         Begin VB.TextBox TxtCustCode 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2520
            TabIndex        =   0
            ToolTipText     =   "Enter Customer Code"
            Top             =   270
            Width           =   1335
         End
         Begin VB.TextBox TxtCustName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2520
            TabIndex        =   3
            ToolTipText     =   "Enter Name Of Customer"
            Top             =   810
            Width           =   4215
         End
         Begin VB.TextBox TxtAdd 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2520
            TabIndex        =   4
            ToolTipText     =   "Enter Custmoer Address"
            Top             =   1350
            Width           =   4215
         End
         Begin VB.TextBox TxtAccName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   4860
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   3330
            Width           =   2055
         End
         Begin VB.TextBox TxtRate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   7020
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   3330
            Width           =   1425
         End
         Begin MSMask.MaskEdBox MaskMobile 
            Height          =   375
            Left            =   4770
            TabIndex        =   8
            ToolTipText     =   "Enter Mobile No"
            Top             =   2790
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "####-#######"
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Date"
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
            Left            =   4230
            TabIndex        =   33
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Service Status"
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
            Left            =   825
            TabIndex        =   32
            Top             =   4950
            Width           =   1575
         End
         Begin VB.Label Label13 
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
            Left            =   1080
            TabIndex        =   31
            Top             =   4500
            Width           =   1320
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
            Left            =   4770
            TabIndex        =   30
            Top             =   4500
            Width           =   1290
         End
         Begin VB.Label Label11 
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
            Left            =   5100
            TabIndex        =   29
            Top             =   3960
            Width           =   945
         End
         Begin VB.Label Label10 
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
            Left            =   7020
            TabIndex        =   28
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label9 
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
            Left            =   1455
            TabIndex        =   27
            Top             =   3960
            Width           =   945
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
            Left            =   480
            TabIndex        =   26
            Top             =   3330
            Width           =   1920
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Phone No / Mobile No"
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
            Left            =   75
            TabIndex        =   25
            Top             =   2790
            Width           =   2325
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
            Left            =   1290
            TabIndex        =   24
            Top             =   360
            Width           =   1110
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
            Left            =   345
            TabIndex        =   23
            Top             =   900
            Width           =   2055
         End
         Begin VB.Label Label4 
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
            Left            =   1515
            TabIndex        =   22
            Top             =   1350
            Width           =   885
         End
         Begin VB.Label Label5 
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
            Left            =   1965
            TabIndex        =   21
            Top             =   2340
            Width           =   435
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   525
      TabIndex        =   16
      Top             =   6660
      Width           =   9015
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Height          =   1275
         Left            =   90
         TabIndex        =   17
         Top             =   90
         Width           =   8745
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
            Left            =   6465
            Style           =   1  'Graphical
            TabIndex        =   43
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
            Left            =   6465
            Style           =   1  'Graphical
            TabIndex        =   42
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
            Left            =   3945
            Style           =   1  'Graphical
            TabIndex        =   41
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
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   40
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
            Left            =   2685
            Style           =   1  'Graphical
            TabIndex        =   39
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
            Left            =   1425
            Style           =   1  'Graphical
            TabIndex        =   38
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
            Left            =   2685
            Style           =   1  'Graphical
            TabIndex        =   37
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
            Left            =   1425
            Style           =   1  'Graphical
            TabIndex        =   36
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
            Left            =   3945
            Style           =   1  'Graphical
            TabIndex        =   35
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
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   270
            Width           =   1275
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Customer Details"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   270
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3240
      Picture         =   "FrmCustTran.frx":0063
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "FrmCustTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCust As New ADODB.Recordset
Dim rsAMast As New ADODB.Recordset
Dim q As String
Dim a As Double
Dim amt As Double
Dim dis As Double
Private Sub CmbAccCode_Change()
If Not (CmbaccCode.Text = "") Then
        rsAMast.Filter = "Acode = " & CmbaccCode.Text
        If Not (rsAMast.EOF) Or Not (rsAMast.BOF) Then
            TxtAccName.Text = IntoStr(rsAMast.Fields(1))
            TxtRate.Text = IntoStr(rsAMast.Fields(4))
         End If
    End If
End Sub
Private Sub CmbAccCode_Click()
If Not (CmbaccCode.Text = "") Then
        rsAMast.Filter = "Acode = " & CmbaccCode.Text
        If Not (rsAMast.EOF) Or Not (rsAMast.BOF) Then
            TxtAccName.Text = IntoStr(rsAMast.Fields(1))
            TxtRate.Text = IntoStr(rsAMast.Fields(4))
                   End If
    End If
End Sub
Private Sub CmdAdd_Click()
Call ClearText(Me)
Frame2.Enabled = True
rsCust.AddNew
End Sub
Private Sub CmdDelete_Click()
rsCust.Delete
Call ClearText(Me)
rsCust.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdEdit_Click()
Frame2.Enabled = True
Call TextToRecordset
rsCust.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsCust.MoveFirst
Call RecordsetToText
End Sub
Private Sub CmdLast_Click()
Frame2.Enabled = False
CmdNext.Enabled = False
rsCust.MoveLast
Call RecordsetToText
End Sub
Private Sub CmdNext_Click()
Frame2.Enabled = False
CmdPrevious.Enabled = True
 rsCust.MoveNext
    If rsCust.EOF = False Then
        Call RecordsetToText
    Else
        rsCust.MoveFirst
        Call RecordsetToText
    End If
End Sub
Private Sub CmdPrevious_Click()
Frame2.Enabled = False
CmdNext.Enabled = True
rsCust.MovePrevious
    If rsCust.BOF = False Then
        Call RecordsetToText
    Else
        rsCust.MoveLast
        Call RecordsetToText
    End If
End Sub



Private Sub CmdPrint_Click()
CustTrans.Show
End Sub

Private Sub CmdSave_Click()
Frame2.Enabled = True
Call TextToRecordset
rsCust.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from CustTran"
rsCust.CursorLocation = adUseClient
rsCust.Open q, cn, adOpenKeyset, adLockOptimistic
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
rsCust.Fields(0) = Val(TxtCustCode.Text)
rsCust.Fields(1) = IntoStr(Trim(TxtCustName.Text))
rsCust.Fields(2) = IntoStr(Trim(TxtAdd.Text))
rsCust.Fields(3) = IntoStr(Trim(TxtAdd1.Text))
rsCust.Fields(4) = IntoStr(Trim(CmbCity.Text))
rsCust.Fields(5) = IntoStr(MaskPhone.Text)
rsCust.Fields(6) = IntoStr(MaskMobile.Text)
rsCust.Fields(7) = IntoStr(Val(CmbaccCode.Text))
rsCust.Fields(8) = IntoStr(Val(TxtQty.Text))
rsCust.Fields(9) = IntoStr(Val(TxtRate.Text))
rsCust.Fields(10) = IntoStr(Val(TxtDis.Text))
rsCust.Fields(11) = IntoStr(Val(TxtNAmt.Text))
rsCust.Fields(12) = IntoStr(Val(TxtCdays.Text))
rsCust.Fields(13) = IntoStr(Val(CmbService.Text))
rsCust.Fields(14) = DTPicker1.Value
End Function
Public Function RecordsetToText()
TxtCustCode.Text = rsCust.Fields(0)
TxtCustName.Text = rsCust.Fields(1)
TxtAdd.Text = rsCust.Fields(2)
TxtAdd1.Text = rsCust.Fields(3)
CmbCity.Text = rsCust.Fields(4)
MaskPhone.Text = rsCust.Fields(5)
MaskMobile.Text = rsCust.Fields(6)
CmbaccCode.Text = rsCust.Fields(7)
TxtQty.Text = rsCust.Fields(8)
TxtRate.Text = rsCust.Fields(9)
TxtDis.Text = rsCust.Fields(10)
TxtNAmt.Text = rsCust.Fields(11)
TxtCdays.Text = rsCust.Fields(12)
CmbService.Text = rsCust.Fields(13)
DTPicker1.Value = rsCust.Fields(14)
End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsAMast = Nothing
Set rsCust = Nothing
End Sub
Private Sub TxtCustName_KeyPress(KeyAscii As Integer)
Call CheckName(KeyAscii)
End Sub
Private Sub TxtDis_Click()
a = InputBox("Enter The Rate Of Discount")
TxtDis.Text = a
amt = TxtRate.Text * TxtQty.Text
dis = amt * a / 100
TxtNAmt.Text = amt - dis
End Sub
