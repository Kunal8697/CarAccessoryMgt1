VERSION 5.00
Begin VB.Form FemReportSelection 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   623
      TabIndex        =   3
      Top             =   1620
      Width           =   5415
      Begin VB.Frame Frame2 
         BackColor       =   &H00EE8EDD&
         Height          =   2055
         Left            =   630
         TabIndex        =   4
         Top             =   180
         Width           =   3975
         Begin VB.CheckBox OptionCurr 
            BackColor       =   &H00EE8EDD&
            Caption         =   "Current Record"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   990
            TabIndex        =   8
            Top             =   900
            Width           =   2175
         End
         Begin VB.CheckBox OptAll 
            BackColor       =   &H00EE8EDD&
            Caption         =   "All Records"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   990
            TabIndex        =   7
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2070
            TabIndex        =   6
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Ok"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   630
            TabIndex        =   5
            Top             =   1440
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1185
      Left            =   630
      TabIndex        =   0
      Top             =   270
      Width           =   5325
      Begin VB.Frame Frame4 
         BackColor       =   &H00EE8EDD&
         Height          =   735
         Left            =   270
         TabIndex        =   1
         Top             =   180
         Width           =   4605
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Report Selection"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   1260
            TabIndex        =   2
            Top             =   270
            Width           =   2025
         End
      End
   End
End
Attribute VB_Name = "FemReportSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
    FrmMdi.CR1.Reset
    FrmMdi.CR1.WindowShowPrintBtn = True
    FrmMdi.CR1.WindowState = crptMaximized
    Screen.MousePointer = vbDefault
    FrmMdi.CR1.WindowShowPrintSetupBtn = True
    FrmMdi.CR1.WindowShowPrintBtn = True
    Select Case (iRptCaller)
    Case 1
        FrmMdi.CR1.WindowTitle = "Accessorie Details"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptAccMast.rpt"
        If OptionCurr.Value = True Then
            FrmMdi.CR1.SelectionFormula = "{AccessaryMaster.Acode} =" & FrmAccMaster.TxtAccCode.Text
        End If
    End Select
    FrmMdi.CR1.DiscardSavedData = True
    FrmMdi.CR1.Action = 1
    Screen.MousePointer = vbDefault
    iRptCaller = 0
    Unload Me
End Sub
