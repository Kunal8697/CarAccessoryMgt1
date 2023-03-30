VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FemServiceBill 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   578
      TabIndex        =   21
      Top             =   4860
      Width           =   7395
      Begin VB.Frame Frame6 
         BackColor       =   &H00EE8EDD&
         Height          =   1275
         Left            =   180
         TabIndex        =   22
         Top             =   90
         Width           =   7035
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
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   270
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   720
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
            TabIndex        =   27
            Top             =   720
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
            Top             =   270
            Width           =   1275
         End
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
            TabIndex        =   23
            Top             =   270
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3525
      Left            =   578
      TabIndex        =   1
      Top             =   1080
      Width           =   7395
      Begin VB.Frame Frame2 
         BackColor       =   &H00EE8EDD&
         Height          =   2985
         Left            =   360
         TabIndex        =   2
         Top             =   180
         Width           =   6585
         Begin VB.TextBox TxtQty 
            Height          =   375
            Left            =   4140
            TabIndex        =   12
            Top             =   1800
            Width           =   1425
         End
         Begin VB.TextBox TxtRate 
            Height          =   375
            Left            =   1530
            TabIndex        =   11
            Top             =   1800
            Width           =   1515
         End
         Begin VB.TextBox TxtAccName 
            Height          =   375
            Left            =   3150
            TabIndex        =   10
            Top             =   1350
            Width           =   3315
         End
         Begin VB.TextBox TxtBillNo 
            Height          =   375
            Left            =   1530
            TabIndex        =   9
            Top             =   450
            Width           =   1335
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
            TabIndex        =   8
            Top             =   900
            Width           =   1545
         End
         Begin VB.TextBox TxtCustName 
            Height          =   375
            Left            =   3150
            TabIndex        =   7
            Top             =   900
            Width           =   3315
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
            TabIndex        =   6
            Top             =   1350
            Width           =   1545
         End
         Begin VB.TextBox TxtDis 
            Height          =   375
            Left            =   1530
            TabIndex        =   4
            Top             =   2250
            Width           =   1515
         End
         Begin VB.TextBox TxtAmt 
            Height          =   375
            Left            =   4140
            TabIndex        =   3
            Top             =   2250
            Width           =   1425
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4050
            TabIndex        =   5
            Top             =   450
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   661
            _Version        =   393216
            Format          =   24641537
            CurrentDate     =   37926
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Discount"
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
            Left            =   450
            TabIndex        =   20
            Top             =   2340
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Quantity"
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
            Left            =   3150
            TabIndex        =   19
            Top             =   1800
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Bill No"
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
            Left            =   630
            TabIndex        =   18
            Top             =   450
            Width           =   705
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Cust Code"
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
            Left            =   285
            TabIndex        =   17
            Top             =   900
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Acc Code"
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
            Left            =   375
            TabIndex        =   16
            Top             =   1350
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Rate"
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
            Left            =   840
            TabIndex        =   15
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Amount"
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
            Left            =   3240
            TabIndex        =   14
            Top             =   2250
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00EE8EDD&
            Caption         =   "Bill Date"
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
            Left            =   2970
            TabIndex        =   13
            Top             =   540
            Width           =   900
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2625
      Picture         =   "FemServiceBill.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00EE8EDD&
      Caption         =   "Service Bill Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      TabIndex        =   0
      Top             =   270
      Width           =   2325
   End
End
Attribute VB_Name = "FemServiceBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
