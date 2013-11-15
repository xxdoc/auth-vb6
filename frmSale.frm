VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBTFSale 
   Caption         =   "Authorize a Sale or Hold"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOldInvoice 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Text            =   "Old Auth"
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame fraTransType 
      Caption         =   "Transaction Type"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   4815
      Begin VB.OptionButton optDepositHold 
         Caption         =   "Deposit Hold"
         Height          =   495
         Left            =   3120
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton optCredit 
         Caption         =   "Credit"
         Height          =   495
         Left            =   1800
         TabIndex        =   17
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton optSale 
         Caption         =   "Sale"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtJob 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Text            =   "Example Auth"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtInvoice 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Text            =   "12345"
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtToken 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "1.00"
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtCustName 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox txtTID 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "38675880-47b0-11e3-8f96-0800200c9a66"
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAuthorize 
      Caption         =   "Authorize Transaction"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label lblConnStatus 
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Label lblOldInvoice 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Invoice"
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblJobCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Job Code"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblInvoice 
      Alignment       =   1  'Right Justify
      Caption         =   "Invoice"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblToken 
      Alignment       =   1  'Right Justify
      Caption         =   "BT Auth Token"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblAmount 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblTransId 
      Alignment       =   1  'Right Justify
      Caption         =   "Transaction ID"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmBTFSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private authType As AuthRequestType

Private Sub cmdAuthorize_Click()
    Dim btTrans As clsBTTransaction
    Dim xmlRequest As String
    Dim strRequestURL As String
    
    Set btTrans = New clsBTTransaction
    
    If txtToken.Text <> "" Then
        If IsNumeric(txtAmount.Text) And CSng(txtAmount.Text) > 0 Then
            btTrans.token = txtToken.Text
            btTrans.amount = txtAmount.Text
            btTrans.custName = txtCustName.Text
            btTrans.invoice = txtInvoice.Text
            btTrans.jobid = txtJob.Text
            btTrans.transid = txtTID.Text
            btTrans.oldInvoice = txtOldInvoice.Text
            
            If optSale.Value = True Then
                btTrans.tType = sale
            ElseIf optCredit.Value = True Then
                btTrans.tType = credit
            ElseIf optDepositHold.Value = True Then
                btTrans.tType = DepositHold
            End If
            
            xmlRequest = CreateXMLRequest(btTrans)
            strRequestURL = modBTF.REQUESTPREFIX
            
            Label1.Caption = strRequestURL
            modBTF.SendAuthPost strRequestURL, xmlRequest, Inet1
            
        Else
            MsgBox "Amount is not a positive number"
        End If
    Else
        MsgBox "Auth request missing token", vbExclamation
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtAmount.Text = "0.00"
    txtCustName.Text = ""
    txtTID = ""
    txtInvoice = ""
    txtJob = ""
    txtToken = ""
    txtOldInvoice = ""
End Sub
Private Sub optCredit_Click()
    txtOldInvoice.Visible = True
    txtOldInvoice.Enabled = True
    lblOldInvoice.Visible = True
End Sub

Private Sub optDepositHold_Click()
    txtOldInvoice.Visible = False
    txtOldInvoice.Enabled = False
    lblOldInvoice.Visible = False
End Sub

Private Sub optSale_Click()
    txtOldInvoice.Visible = False
    txtOldInvoice.Enabled = False
    lblOldInvoice.Visible = False
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Dim strChunk As String
    Dim strResponse As String
    
    strResponse = ""
    Select Case State
    Case icError
        MsgBox Inet1.ResponseInfo, , "Connection Error"
        lblConnStatus.Caption = "ERROR Connecting" & Inet1.ResponseCode
    Case icResolvingHost
        lblConnStatus.Caption = "Resolving host..."
     Case icHostResolved
        lblConnStatus.Caption = "Host resolved..."
    Case icConnecting
        lblConnStatus.Caption = "Connecting..."
    Case icConnected
        lblConnStatus.Caption = "Connected..."
    Case icRequesting
        lblConnStatus.Caption = "Sending request..."
    Case icReceivingResponse
        lblConnStatus.Caption = "Receiving response..."
    Case icResponseReceived
        lblConnStatus.Caption = "Response received..."
    Case icResponseCompleted
        lblConnStatus.Caption = "Response Complete"
        strChunk = Inet1.GetChunk(1024, icString)
        Do While Len(strChunk) > 0
            strResponse = strResponse & strChunk
            strChunk = Inet1.GetChunk(1024, icString)
        Loop
        
        strRespHeader = Inet1.GetHeader
        If InStr(1, strRespHeader, "200") Then
            Label1.Caption = modBTF.ParseAuthResponse(strResponse)
        Else
            MsgBox strRespHeader
            strResponse = "REQUEST ERROR"
        End If
    Case icDisconnecting
        lblConnStatus.Caption = "Disconnecting..."
    Case icDisconnected
        lblConnStatus.Caption = "Disconnected."
    Case Else
        lblConnStatus.Caption = "Unknown State: " & State
    End Select
End Sub

Public Sub setCustName(ByVal strName As String)
    txtCustName.Text = strName
End Sub

Public Sub setToken(ByVal strToken As String)
    txtToken.Text = strToken
End Sub
