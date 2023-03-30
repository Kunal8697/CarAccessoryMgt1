VERSION 5.00
Begin VB.Form FrmPhotoAll 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11250
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF80FF&
      Height          =   5535
      Left            =   1920
      ScaleHeight     =   5475
      ScaleWidth      =   8115
      TabIndex        =   3
      Top             =   1080
      Width           =   8175
   End
   Begin VB.ComboBox CmbCategory 
      Height          =   315
      ItemData        =   "FrmPhotoAll.frx":0000
      Left            =   1560
      List            =   "FrmPhotoAll.frx":0002
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00FF80FF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6870
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   90
      Top             =   2340
   End
   Begin VB.CommandButton CmdPreview 
      BackColor       =   &H00FF80FF&
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5250
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6870
      Width           =   1200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   90
      Top             =   2940
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Decoration Accessories"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   600
      Width           =   3045
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5760
      TabIndex        =   7
      Top             =   540
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   450
      TabIndex        =   6
      Top             =   540
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "File Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4500
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "0 of 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7860
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "FrmPhotoAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As String
Dim i As Integer
Dim iMax As Integer
Dim p As IPictureDisp
Dim j As Integer
Private Sub CmbCategory_Click()
i = 0
    iMax = CmbCategory.ItemData(CmbCategory.ListIndex)
End Sub
Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub CmdPreview_Click()
If CmdPreview.Caption = "Preview" Then
        CmdPreview.Caption = "Stop"
        Timer1.Enabled = True
    Else
        CmdPreview.Caption = "Preview"
        Timer1.Enabled = False
        Timer2.Enabled = False
        Picture1.PaintPicture p, 0, 0
    End If
End Sub

Private Sub Form_Load()
Call CenterInScreen(Me)
Call FillCombo
    s = App.Path & "\Album\Decoration\"
    i = 1
    Set p = LoadPicture(s & CmbCategory.Text & i & ".jpg")
    Label3.Caption = CmbCategory.Text & i & ".jpg"
    'Call CenterInScreen(Me)
    j = Picture1.Width
    CmdPreview.Caption = "Stop"
    Call ShowNext
End Sub
Private Sub ShowNext()
    If i < iMax And i >= 0 Then
        i = i + 1
        Set p = LoadPicture(s & CmbCategory.Text & i & ".jpg")
        Label3.Caption = CmbCategory.Text & i & ".jpg"
        Label5.Caption = i & " of " & iMax
        Picture1.Cls
        Timer2.Enabled = True
    End If
    If i = iMax Then
        i = 0
    End If
End Sub
Private Sub FillCombo()
Dim rsTemp As New ADODB.Recordset
Dim j As Integer
    rsTemp.Open "photo_category", cn, adOpenForwardOnly, adLockReadOnly
    j = 0
    While rsTemp.EOF = False
        CmbCategory.AddItem IntoStr(rsTemp.Fields(1))
        CmbCategory.ItemData(j) = Val(IntoStr(rsTemp.Fields(2)))
        rsTemp.MoveNext
        j = j + 1
    Wend
    If j > 0 Then
        CmbCategory.ListIndex = 0
    End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
    Call ShowNext
End Sub

Private Sub Timer2_Timer()
  Picture1.PaintPicture p, j, 0
    j = j - 100
    If j <= 0 Then
        j = Picture1.Width
        Timer2.Enabled = False
        Timer1.Enabled = True
        Picture1.PaintPicture p, 0, 0
    End If
End Sub
