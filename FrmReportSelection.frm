VERSION 5.00
Begin VB.Form FrmReportSelection 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Report Selection"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1305
      Left            =   600
      TabIndex        =   6
      Top             =   240
      Width           =   5385
      Begin VB.Frame Frame4 
         BackColor       =   &H00EE8EDD&
         Height          =   735
         Left            =   360
         TabIndex        =   7
         Top             =   240
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
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   2025
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EE8EDD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   5415
      Begin VB.Frame Frame2 
         BackColor       =   &H00EE8EDD&
         Height          =   2055
         Left            =   630
         TabIndex        =   1
         Top             =   180
         Width           =   4335
         Begin VB.OptionButton OptionCurr 
            BackColor       =   &H00EE8EDD&
            Caption         =   "Current Record's"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   5
            Top             =   840
            Width           =   2415
         End
         Begin VB.OptionButton OptAll 
            BackColor       =   &H00EE8EDD&
            Caption         =   "All Record's"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   4
            Top             =   360
            Width           =   2535
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
            TabIndex        =   3
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
            Left            =   720
            TabIndex        =   2
            Top             =   1440
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "FrmReportSelection"
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
    Case 2
        FrmMdi.CR1.WindowTitle = "Service Details"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptServiceMast.rpt"
        If OptionCurr.Value = True Then
            FrmMdi.CR1.SelectionFormula = "{ServiceMaster.Scode} = " & FrmSMast.TxtSCode.Text
        End If
    Case 3
        FrmMdi.CR1.WindowTitle = "Employee Details"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptEmpMast.rpt"
        If OptionCurr.Value = True Then
            FrmMdi.CR1.SelectionFormula = "{EmpMaster.EmpCode} = " & FrmEmp.TxtECode.Text
        End If
     Case 4
        FrmMdi.CR1.WindowTitle = "Purchases Details"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptPOrder.rpt"
        If OptionCurr.Value = True Then
            FrmMdi.CR1.SelectionFormula = "{PurchaseTran.PoNo} = " & FrmPOrder.TxtPno.Text
        End If
    Case 5
        FrmMdi.CR1.WindowTitle = "Stock Keeping"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptStockTran.rpt"
        If OptionCurr.Value = True Then
            FrmMdi.CR1.SelectionFormula = "{StockTran.Acode} = " & FrmSTrans.CmbAcode.Text
        End If
    Case 6
        FrmMdi.CR1.WindowTitle = "Customer Details"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptCustTrans.rpt"
        If OptionCurr.Value = True Then
            FrmMdi.CR1.SelectionFormula = "{CustTran.CustCode} = " & FrmCustTran.TxtCustCode.Text
        End If
        Case 7
        'Service Details
        FrmMdi.CR1.WindowTitle = "Service Details"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptSTrans.rpt"
        If OptionCurr.Value = True Then
            FrmMdi.CR1.SelectionFormula = "{SStatusTran.CustCode} = " & FrmSStatus.CmbCustCode
        End If
          Case 8
        
        FrmMdi.CR1.WindowTitle = "Employee Monthly Details"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptEmpTrans.rpt"
        If OptionCurr.Value = True Then
            FrmMdi.CR1.SelectionFormula = "{EmpTran.Ecode} = " & FrmEmpTran.CmbEmpCode.Text
        End If
        Case 9
        FrmMdi.CR1.WindowTitle = "Quatation"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptQuatation.rpt"
        If OptionCurr.Value = True Then
        FrmMdi.CR1.SelectionFormula = "{Quotation.SrNo} =" & FrmQuatation.TxtSrNo.Text
        End If
    
    Case 10
        FrmMdi.CR1.WindowTitle = "Bill Details"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptAccBill.rpt"
        If OptionCurr.Value = True Then
        FrmMdi.CR1.SelectionFormula = "{ABillTran.BCode} =" & FrmAccBill.TxtBillNo.Text
        End If
     
    Case 11
        FrmMdi.CR1.WindowTitle = "Service Bill Detals"
        FrmMdi.CR1.ReportFileName = "C:\CarAccessoryMgt\RptServiceBill.rpt"
        If OptionCurr.Value = True Then
        FrmMdi.CR1.SelectionFormula = "{SBillTran.BCode} =" & FrmServiceBill.TxtBillNo.Text
        End If
      
        End Select
    FrmMdi.CR1.DiscardSavedData = True
    FrmMdi.CR1.Action = 1
    Screen.MousePointer = vbDefault
    iRptCaller = 0
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me

End Sub
