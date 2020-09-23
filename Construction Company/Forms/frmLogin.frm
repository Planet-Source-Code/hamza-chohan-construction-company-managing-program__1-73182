VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4095
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1781.363
   ScaleMode       =   0  'User
   ScaleWidth      =   3844.983
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc ADOUser 
      Height          =   330
      Left            =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Database\Data.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Database\Data.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SysLogin"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "frmLogin.frx":08CA
      ScaleHeight     =   855
      ScaleWidth      =   3015
      TabIndex        =   5
      Top             =   120
      Width           =   3015
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "System Login"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   765
      TabIndex        =   3
      Top             =   2220
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2370
      TabIndex        =   4
      Top             =   2220
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      DataSource      =   "datPrimaryRS"
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1725
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   1680
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Username As String
Public Password As String
Public Login As Boolean

Private Sub cmdCancel_Click()
    End

End Sub

Private Sub cmdOK_Click()
If Login = True Then
    If cboUser.Text = Username And txtPassword.Text = Password Then
        txtPassword.Text = ""
        cboUser.SetFocus
        Login = False
    Else
        MsgBox "Sorry! Username or Password is Wrong", vbCritical + vbOKOnly, "Login Error"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
Else
    ADOUser.Recordset.Filter = "Username = '" & cboUser.Text & "'"
        If txtPassword = ADOUser.Recordset!Password Then
            Username = cboUser.Text
            Password = txtPassword.Text
            Username = StrConv(ADOUser.Recordset!Username, vbUpperCase)
            Unload Me
            frmMain.Show
        Else
            MsgBox "Sorry! Username or Password is Wrong", vbCritical + vbOKOnly, "Login Error"
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
        End If
End If
End Sub

Private Sub Form_Load()
ADOUser.Refresh
Do While Not ADOUser.Recordset.EOF
    cboUser.AddItem ADOUser.Recordset!Username
    ADOUser.Recordset.MoveNext
Loop
End Sub
