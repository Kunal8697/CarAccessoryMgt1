VERSION 5.00
Begin VB.Form FrmSTran 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CmbAcode 
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
      Left            =   2700
      TabIndex        =   18
      ToolTipText     =   "Select Acc Code"
      Top             =   2070
      Width           =   1365
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   308
      TabIndex        =   4
      Top             =   4590
      Width           =   7125
      Begin VB.Frame Frame6 
         BackColor       =   &H00EE8EDD&
         Height          =   1275
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   6945
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
            TabIndex        =   15
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
            TabIndex        =   12
            Top             =   270
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   720
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   270
            Width           =   1275
         End
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
            TabIndex        =   6
            Top             =   720
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   308
      TabIndex        =   0
      Top             =   1170
      Width           =   7125
      Begin VB.Frame Frame2 
         BackColor       =   &H00EE8EDD&
         Height          =   2175
         Left            =   720
         TabIndex        =   1
         Top             =   450
         Width           =   5595
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   3150
            TabIndex        =   19
            Top             =   450
            Width           =   2265
         End
         Begin VB.TextBox TxtQty 
            Height          =   375
            Left            =   1710
            TabIndex        =   2
            ToolTipText     =   "Enter Quantity"
            Top             =   1170
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
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
            Left            =   720
            TabIndex        =   17
            Top             =   1170
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "A code"
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
            TabIndex        =   3
            Top             =   450
            Width           =   735
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00EE8EDD&
      Caption         =   "Stock Details"
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
      Left            =   2790
      TabIndex        =   16
      Top             =   540
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1800
      Picture         =   "FrmSTran.frx":0000
      Top             =   450
      Width           =   480
   End
End
Attribute VB_Name = "FrmStran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsStock As New ADODB.Recordset
Dim rsAMast As New ADODB.Recordset
Dim q As String
Private Sub CmbAcode_Change()
If Not (CmbAcode.Text = "") Then
        rsAMast.Filter = "Acode = " & CmbAcode.Text
        If Not (rsAMast.EOF) Or Not (rsAMast.BOF) Then
            Text1.Text = IntoStr(rsAMast.Fields(1))
            
        End If
    End If
End Sub
Private Sub CmbAcode_Click()
If Not (CmbAcode.Text = "") Then
        rsAMast.Filter = "Acode = " & CmbAcode.Text
        If Not (rsAMast.EOF) Or Not (rsAMast.BOF) Then
            Text1.Text = IntoStr(rsAMast.Fields(1))
            
        End If
    End If
End Sub
Private Sub CmdAdd_Click()
rsStock.AddNew
End Sub

Private Sub CmdDelete_Click()
rsStock.Delete
Call ClearText(Me)
rsStock.MoveFirst
Call RecordsetToText

End Sub

Private Sub CmdEdit_Click()
Call TextToRecordset
rsStock.Update

End Sub

Private Sub CmdExit_Click()
Unload Me

End Sub

Private Sub CmdFirst_Click()
CmdPrevious.Enabled = False
CmdNext.Enabled = True
rsStock.MoveFirst
Call RecordsetToText

End Sub

Private Sub CmdLast_Click()
CmdNext.Enabled = False
rsStock.MoveLast
Call RecordsetToText

End Sub

Private Sub CmdNext_Click()
CmdPrevious.Enabled = True
 rsStock.MoveNext
    If rsStock.EOF = False Then
        Call RecordsetToText
    Else
        rsStock.MoveFirst
        Call RecordsetToText
    End If

End Sub

Private Sub CmdPrevious_Click()
CmdNext.Enabled = True
rsStock.MovePrevious

    If rsStock.BOF = False Then
        Call RecordsetToText
    Else
        rsStock.MoveLast
        Call RecordsetToText
    End If

End Sub

Private Sub CmdPrint_Click()
iRptCaller = 5
FrmReportSelection.Show
End Sub

Private Sub CmdSave_Click()
Call TextToRecordset
rsStock.Update
Call ClearText(Me)

End Sub

Private Sub Form_Load()
Call CenterInScreen(Me)
q = "select * from StockTran"
rsStock.CursorLocation = adUseClient
rsStock.Open q, cn, adOpenKeyset, adLockOptimistic
rsAMast.Open "AccessaryMaster", cn, adOpenKeyset, adLockOptimistic
Call FillComboMenu
End Sub

Public Function TextToRecordset()
rsStock.Fields(0) = Val(CmbAcode.Text)
rsStock.Fields(1) = Val(TxtQty.Text)

End Function

Public Function RecordsetToText()
CmbAcode.Text = rsStock.Fields(0)
TxtQty.Text = rsStock.Fields(1)

End Function
Private Sub FillComboMenu()

Dim j As Integer
   
    j = 0
    While rsAMast.EOF = False
        CmbAcode.AddItem IntoStr(rsAMast.Fields(0))
        rsAMast.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbAcode.ListIndex = 0
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsStock = Nothing
Set rsAMast = Nothing

End Sub

