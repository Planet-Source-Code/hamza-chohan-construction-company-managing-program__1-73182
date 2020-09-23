VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOrders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orders"
   ClientHeight    =   5805
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "frmOrders.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLocation 
      DataField       =   "LocationType"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      ItemData        =   "frmOrders.frx":08CA
      Left            =   2760
      List            =   "frmOrders.frx":08D4
      TabIndex        =   20
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   5040
      TabIndex        =   19
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   420
      Left            =   3840
      TabIndex        =   18
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   420
      Left            =   2640
      TabIndex        =   17
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   420
      Left            =   1440
      TabIndex        =   16
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   420
      Left            =   240
      TabIndex        =   15
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "ServiceNeeded"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2760
      TabIndex        =   13
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "LocationCity"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   11
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "LocationAddress"
      DataSource      =   "datPrimaryRS"
      Height          =   645
      Index           =   4
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "DateOrderCompleted"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "DateOrderPlaced"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "CustomerId"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "OrderId"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   240
      Top             =   5280
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      RecordSource    =   "select OrderId,CustomerId,DateOrderPlaced,DateOrderCompleted,LocationAddress,LocationCity,ServiceNeeded,LocationType from Orders"
      Caption         =   " "
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
   Begin VB.Label lblLabels 
      Caption         =   "LocationType:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ServiceNeeded:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LocationCity:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LocationAddress:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DateOrderCompleted:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Caption         =   "DateOrderPlaced:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   "CustomerId:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "OrderId:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
    If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete") = vbYes Then
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End If
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

