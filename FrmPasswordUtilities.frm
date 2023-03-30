VERSION 5.00
Begin VB.Form FrmPasswordUtilities 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7305
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1605
      Left            =   360
      TabIndex        =   7
      Top             =   3780
      Width           =   6765
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         Height          =   1095
         Left            =   540
         TabIndex        =   8
         Top             =   180
         Width           =   5535
         Begin VB.CommandButton CmdExit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmdSave 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmdAdd 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2445
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   6675
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   1755
         Left            =   450
         TabIndex        =   2
         Top             =   270
         Width           =   5745
         Begin VB.TextBox txtPass 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "*"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   2595
            PasswordChar    =   "*"
            TabIndex        =   4
            ToolTipText     =   "Enter Password"
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox TxtUName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2595
            TabIndex        =   3
            ToolTipText     =   "Enter User Name"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "User Name"
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
            TabIndex        =   6
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "PassWord"
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
            Left            =   1305
            TabIndex        =   5
            Top             =   840
            Width           =   1080
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1200
      Picture         =   "FrmPasswordUtilities.frx":0000
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Password Confirmation"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   3075
   End
End
Attribute VB_Name = "FrmPasswordUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUser As New ADODB.Recordset
Private Sub CmdAdd_Click()
Call ClearText(Me)
rsUser.AddNew
End Sub
Private Sub CmdEdit_Click()
Call TextToRecordset
rsUser.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub


Private Sub CmdSave_Click()
Call TextToRecordset
rsUser.Update
'Call ClearText(Me)
MsgBox "Entry Accepted For New Entry Press Add Button", vbOKOnly, "Car Accessories Management System"
End Sub
Private Sub Form_Load()
Call CenterInScreen(Me)
rsUser.CursorLocation = adUseClient
rsUser.Open "Select * from UserMaster ", cn, adOpenKeyset, adLockOptimistic
End Sub
Public Function TextToRecordset()
rsUser.Fields(0) = Trim(TxtUName.Text)
rsUser.Fields(1) = Trim(txtPass.Text)
End Function
Public Function RecordsetToText()
TxtUName.Text = IntoStr(rsUser.Fields(0))
txtPass.Text = IntoStr(rsUser.Fields(1))
End Function
Private Sub Form_Unload(Cancel As Integer)
Set rsUser = Nothing
End Sub
