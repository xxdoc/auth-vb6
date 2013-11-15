VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBTFDepositCollect 
   Caption         =   "Deposit Collect"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox TransList 
      Columns         =   1
      Height          =   2790
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   7335
   End
   Begin VB.TextBox txtAuthSeq 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtTID 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Text            =   "38675880-47b0-11e3-8f96-0800200c9a66"
      Top             =   0
      Width           =   3135
   End
   Begin VB.TextBox txtCustName 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "1.00"
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtToken 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtInvoice 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtJob 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton cmdCollectTransaction 
      Caption         =   "Collect"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdListCollectables 
      Caption         =   "List Transactions"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblAuthSeq 
      Caption         =   "Auth Sequence"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblTransId 
      Alignment       =   1  'Right Justify
      Caption         =   "Transaction ID"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblAmount 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblToken 
      Alignment       =   1  'Right Justify
      Caption         =   "BT Auth Token"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblInvoice 
      Alignment       =   1  'Right Justify
      Caption         =   "Invoice"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblJobCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Job Code"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label lblConnStatus 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "frmBTFDepositCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private transArray() As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCollectTransaction_Click()
    Dim strRequestURL As String
    Dim strRequestXML As String
    Dim btTrans As clsBTTransaction
    
    Set btTrans = New clsBTTransaction
    
    If txtAuthSeq.Text <> "" _
    And txtToken.Text <> "" Then
        btTrans.tType = DepositCollect
        btTrans.authseq = txtAuthSeq.Text
        btTrans.token = txtToken.Text
        btTrans.amount = txtAmount.Text
        btTrans.custName = txtCustName.Text
        btTrans.invoice = txtInvoice.Text
        btTrans.jobid = txtJob.Text
        btTrans.transid = txtTID.Text
        strRequestURL = modBTF.REQUESTPREFIX
        strRequestXML = modBTF.CreateXMLRequest(btTrans)
        
        Label1.Caption = strRequestURL
        modBTF.SendAuthPost strRequestURL, strRequestXML, Inet1
        TransList.Clear
    Else
        MsgBox "Missing Auth sequence or Auth token"
    End If
    
    
End Sub

Private Sub cmdListCollectables_Click()
    Dim strRequestURL As String
    
    TransList.Clear
    
    strRequestURL = modBTF.REQUESTPREFIX + "transactions/deposit"
    Label1.Caption = strRequestURL
    modBTF.SendRequest strRequestURL, Inet1
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Dim strChunk As String
    
    Select Case State
    Case icError
        MsgBox Inet1.ResponseInfo, , "Connection Error"
        lblConnStatus.Caption = "ERROR Connecting" & Inet1.ResponseCode
    Case icResolvingHost
        lblConnStatus.Caption = "Resolving host..."
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
            parseResponse (strResponse)
        Else
            MsgBox strRespHeader
            strResponse = "REQUEST ERROR"
        End If
    End Select
End Sub

Private Sub parseResponse(strResponse As String)
    Dim doc As New MSXML.DOMDocument
    Dim success As Boolean
    Dim strTS As String
    
    success = doc.loadXML(strResponse)
    
    If success Then
        Dim root As MSXML.IXMLDOMNode
        Dim subRoot As MSXML.IXMLDOMNode
        
        Set root = doc.selectSingleNode("bt:bluetarp-authorization")
        Set subRoot = root.selectSingleNode("bt:transaction-response")
        
        If Not subRoot Is Nothing Then
            Dim nodeList As MSXML.IXMLDOMNodeList
            
            Set nodeList = doc.selectNodes("//bt:transactions/*")
            If Not nodeList Is Nothing And nodeList.length > 0 Then
               transArray = ParseTransList(strResponse)
                Dim i As Long
                Dim strEntry As String
                For j = 0 To UBound(transArray, 2) - 1
                    If transArray(0, j) <> "" Then
                        strEntry = transArray(0, j)
                        For i = 1 To UBound(transArray, 1) - 1
                            strEntry = strEntry & "    " & transArray(i, j)
                        Next
                        TransList.AddItem (strEntry)
                    End If
                Next
            Else
                Label1.Caption = "No matches found"
            End If
        Else
            Set subRoot = root.selectSingleNode("bt:authorization-response")
            If Not subRoot Is Nothing Then
                Label1.Caption = ParseAuthResponse(strResponse)
            Else
                Label1.Caption = "Reply not recognized"
            End If
        End If
    Else
        Label1.Caption = "Unable to parse response"
    End If
        
        
End Sub

Private Sub TransList_Click()
    If TransList.ListIndex <> -1 Then
        Dim index As Long
        index = TransList.ListIndex
        
        txtAuthSeq.Text = transArray(0, index)
        txtAmount.Text = transArray(1, index)
        txtToken.Text = transArray(2, index)
        txtCustName.Text = transArray(3, index)
    End If
End Sub
