VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   6465
   ClientLeft      =   2835
   ClientTop       =   3525
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3819.735
   ScaleMode       =   0  'User
   ScaleWidth      =   7154.768
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   2535
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Width           =   4425
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   2085
         Left            =   360
         TabIndex        =   1
         Top             =   180
         Width           =   3735
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   1650
            PasswordChar    =   "*"
            TabIndex        =   5
            Text            =   "Admin"
            ToolTipText     =   "Enter Password"
            Top             =   885
            Width           =   1605
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00E274C1&
            Cancel          =   -1  'True
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1440
            Width           =   1140
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H00E274C1&
            Caption         =   "OK"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1440
            Width           =   1140
         End
         Begin VB.TextBox txtUserName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1680
            TabIndex        =   2
            Text            =   "Admin"
            ToolTipText     =   "Enter User Name"
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00FFFFC0&
            Caption         =   "&Password"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   480
            TabIndex        =   7
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00FFFFC0&
            Caption         =   "&User Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1320
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1320
      Picture         =   "frmLogin.frx":0000
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "  Please Enter Login Password . . . "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   1560
      Width           =   3480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUser As New ADODB.Recordset
Dim Flage As Boolean
Dim q As String
Private Sub cmdCancel_Click()
 Unload Me
End Sub
Private Sub cmdOK_Click()
'On Error GoTo ErrHandler
 rsUser.MoveFirst
    If rsUser.EOF = False And rsUser.BOF = False Then
        While rsUser.EOF = False
            If rsUser.Fields(0) = txtUserName.Text And txtPassword.Text = rsUser.Fields(1) Then
               FrmMdi.Show
                                Unload frmLogin
                Exit Sub
            End If
            rsUser.MoveNext
        Wend
                If Flage = False Then
            MsgBox "Please Enter Valid User Name AND Password, try again!", , "Login Message Box"
            txtPassword.Text = ""
            txtUserName.Text = ""
        End If
    End If
    Exit Sub
ErrHandler:
    Dsc = Err.Number & "  " & Err.Description
    LogError (Dsc)
    MsgBox Dsc, , "CarAccssoriesSys"
End Sub
Private Sub Form_Load()
  
'On Error GoTo ErrHandler
    Call Con
    'Set rsUser = New Recordset
    q = "select * from UserMaster"
    rsUser.CursorLocation = adUseClient
    rsUser.Open q, cn, adOpenKeyset, adLockOptimistic
    Exit Sub
'ErrHandler:
'    Dsc = Err.Number & "  " & Err.Description
'    LogError (Dsc)
'    MsgBox Dsc, , "CarAccssoriesSys"
   End Sub
Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ErrHandler
    Set rsUser = Nothing
    Exit Sub
'ErrHandler:
'    Dsc = Err.Number & "  " & Err.Description
'    LogError (Dsc)
'    MsgBox Dsc, , "CarAccssoriesSys"
End Sub

