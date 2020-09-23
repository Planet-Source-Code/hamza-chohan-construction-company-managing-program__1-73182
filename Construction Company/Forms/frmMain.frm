VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Construction Company Managing Program"
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":08CA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   0
      Top             =   6360
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6915
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A06B
            Text            =   "Time:"
            TextSave        =   "Time:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A945
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "Database"
      Begin VB.Menu mnuAttendence 
         Caption         =   "Attendence"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "Customers"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployees 
         Caption         =   "Employees"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrders 
         Caption         =   "Orders"
         Begin VB.Menu mnuCustomerOrders 
            Caption         =   "Customer Orders"
         End
         Begin VB.Menu sep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEmployeesOrders 
            Caption         =   "Employees Orders"
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuReportCustomer 
         Caption         =   "Customers Report"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomerOrderRpt 
         Caption         =   "Customers Orders Report"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployeesReport 
         Caption         =   "Employees Report"
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdersRpt 
         Caption         =   "Orders Report"
      End
   End
   Begin VB.Menu mnuSecurity 
      Caption         =   "Security"
      Begin VB.Menu mnuUsers 
         Caption         =   "System Users"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuSplash 
         Caption         =   "Show Splash"
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Program"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
StatusBar.Panels(2) = Timer
StatusBar.Panels(4) = Date
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you sure to exit the system?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbYes Then
End
Else
Cancel = True
Exit Sub
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuAttendence_Click()
frmAttendence.Show 1
End Sub

Private Sub mnuCustomerOrderRpt_Click()
rptCustomersOrders.Show
End Sub


Private Sub mnuCustomerOrders_Click()
frmCustomerOrders.Show 1
End Sub

Private Sub mnuCustomers_Click()
frmCustomers.Show 1
End Sub

Private Sub mnuEmployees_Click()
frmEmployees.Show 1
End Sub

Private Sub mnuEmployeesOrders_Click()
frmEmployeeOrders.Show 1
End Sub

Private Sub mnuEmployeesReport_Click()
rptEmployees.Show
End Sub

Private Sub mnuExit_Click()
If MsgBox("Are you sure to exit the system?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbYes Then
End
Else
Cancel = True
Exit Sub
End If
End Sub

Private Sub mnuOrdersRpt_Click()
rptOrders.Show
End Sub

Private Sub mnuReportCustomer_Click()
rptCustomers.Show
End Sub

Private Sub mnuSplash_Click()
frmSplash.Show 1
End Sub

Private Sub mnuUsers_Click()
frmSystemUsers.Show
End Sub

Private Sub Timer_Timer()
StatusBar.Panels(2) = Time
End Sub

