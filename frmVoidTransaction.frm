VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBTFVoidTransaction 
   Caption         =   "Void a transaction"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox TransList 
      Columns         =   1
      Height          =   3960
      Left            =   0
      TabIndex        =   13
      Top             =   2400
      Width           =   6255
   End
   Begin VB.TextBox txtAuthToken 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox txtTransID 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtCustName 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox txtAuthSeq 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdListVoidables 
      Caption         =   "List Transactions"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdVoidTransaction 
      Caption         =   "Void"
      Height          =   375
      Left            =   -120
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblConnStatus 
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblAuthToken 
      Alignment       =   1  'Right Justify
      Caption         =   "Auth Token"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblAuth 
      Alignment       =   1  'Right Justify
      Caption         =   "Auth Sequence"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblTrans 
      Alignment       =   1  'Right Justify
      Caption         =   "Transaction ID"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   4695
   End
End
Attribute VB_Name = "frmBTFVoidTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private transArray() As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdListVoidables_Click()
    Dim strRequestURL As String
    
    TransList.Clear
    
    strRequestURL = modBTF.REQUESTPREFIX + "transactions/void"
    Label1.Caption = strRequestURL
    modBTF.SendRequest strRequestURL, Inet1
    
End Sub

Private Sub cmdVoidTransaction_Click()
    Dim strRequestURL As String
    Dim strRequestXML As String
    Dim btTrans As clsBTTransaction
    
    Set btTrans = New clsBTTransaction
    
    If txtAuthSeq.Text <> "" _
    And txtAuthToken.Text <> "" Then
        btTrans.tType = Void
        btTrans.authseq = txtAuthSeq.Text
        btTrans.token = txtAuthToken.Text
        strRequestURL = modBTF.REQUESTPREFIX
        strRequestXML = modBTF.CreateXMLRequest(btTrans)
        
        Label1.Caption = strRequestURL
        modBTF.SendAuthPost strRequestURL, strRequestXML, Inet1
        TransList.Clear
    Else
        MsgBox "Missing Auth sequence or Auth token"
    End If
    
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
        txtAuthToken.Text = transArray(2, index)
        txtCustName.Text = transArray(3, index)
    End If
End Sub
