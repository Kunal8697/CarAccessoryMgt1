VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FrmMdi 
   BackColor       =   &H00404040&
   Caption         =   "Car Accessories Management System"
   ClientHeight    =   5790
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   6780
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5175
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12-09-2018"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "18:40"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MenuMaster 
      Caption         =   "&Master"
      Begin VB.Menu MastAccessories 
         Caption         =   "Accessories Master"
      End
      Begin VB.Menu MastServiceMaster 
         Caption         =   "Service Master"
      End
      Begin VB.Menu MastEmployeeMaster 
         Caption         =   "Employee Master"
      End
   End
   Begin VB.Menu TranTransactionEntry 
      Caption         =   "&Transaction Entry"
      Begin VB.Menu TranPurchaseOrder 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu TranCustomerDetails 
         Caption         =   "Customer Details"
      End
      Begin VB.Menu TranServiceStatus 
         Caption         =   "Service Status"
      End
      Begin VB.Menu TranAccBillDetail 
         Caption         =   "Acc Bill Detail"
      End
      Begin VB.Menu TranServiceBillStatus 
         Caption         =   "Service Bill Status"
      End
      Begin VB.Menu TranEmployeeDetails 
         Caption         =   "Employee Details"
      End
   End
   Begin VB.Menu TranQuotation 
      Caption         =   "&Quotation"
   End
   Begin VB.Menu MnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnupurorddet 
         Caption         =   "Purchase Order Details"
      End
      Begin VB.Menu mnucustDetails 
         Caption         =   "Customer Details"
      End
      Begin VB.Menu mnuaccbill 
         Caption         =   "Accessories Bill"
      End
      Begin VB.Menu mnuservicebill 
         Caption         =   "Service Bill Details"
      End
      Begin VB.Menu mnuemppayslip 
         Caption         =   "Employee Pay Slip"
      End
      Begin VB.Menu mnuquatation 
         Caption         =   "Quatation"
      End
   End
   Begin VB.Menu MnuPhoto 
      Caption         =   "Photo Allbum"
      Begin VB.Menu mnuDeco 
         Caption         =   "Decoration"
      End
   End
   Begin VB.Menu MenuUtilities 
      Caption         =   "&Utilities"
      WindowList      =   -1  'True
      Begin VB.Menu mnupassword 
         Caption         =   "Password Utilities"
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "FrmMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim N As Double
Dim Cal As Double
Private Sub Exit_Click()
Unload Me
End Sub
Private Sub MastAccessories_Click()
FrmAccMaster.Show
End Sub
Private Sub MastEmployeeMaster_Click()
FrmEmp.Show
End Sub
Private Sub MastServiceMaster_Click()
FrmSMast.Show
End Sub
Private Sub MenuAboutUs_Click()
frmAbout.Show
End Sub
Private Sub MenuCalculator_Click()
Cal = Shell("C:\WINDOWS\system32\CALC.EXE", vbMaximizedFocus)
End Sub
Private Sub MenuHelp1_Click(Index As Integer)
Call Shell("C:\Program Files\Internet Explorer\Iexplore.exe " & App.Path & "\Help\Index.htm", vbMaximizedFocus)
End Sub
Private Sub MenuNotepad_Click()
N = Shell("C:\WINDOWS\NOTEPAD.EXE", vbMaximizedFocus)
End Sub

Private Sub mnuaccbill_Click()
AccBill.Show
End Sub

Private Sub mnucustDetails_Click()
CustTrans.Show
End Sub

Private Sub mnuDeco_Click()
FrmPhotoAll.Show
End Sub

Private Sub mnudoor_Click()
'FrmDoorAcc.Show
End Sub

Private Sub mnuemppayslip_Click()
EmpTrans.Show
End Sub

Private Sub mnufront_Click()
'FrmFront.Show
End Sub

Private Sub mnupassword_Click()
FrmPasswordUtilities.Show
End Sub
Private Sub mnupurorddet_Click()
PurTran.Show
End Sub

Private Sub mnuquatation_Click()
Quat.Show
End Sub

Private Sub mnuservicebill_Click()
SerBill.Show
End Sub

Private Sub mnuside_Click()
'FrmSide.Show
End Sub

Private Sub TranAccBillDetail_Click()
FrmAccBill.Show
End Sub
Private Sub TranCustomerDetails_Click()
FrmCustTran.Show
End Sub
Private Sub TranEmployeeDetails_Click()
FrmEmpTran.Show
End Sub
Private Sub TranPurchaseOrder_Click()
FrmPOrder.Show
End Sub
Private Sub TranQuotation_Click()
FrmQuatation.Show
End Sub
Private Sub TranServiceBillStatus_Click()
FrmServiceBill.Show
End Sub
Private Sub TranServiceStatus_Click()
FrmSStatus.Show
End Sub
Private Sub TranStockKeeping_Click()
'FrmSTrans.Show
End Sub
