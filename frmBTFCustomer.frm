VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBTFCustomer 
   Caption         =   "Lookup a purchaser token"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Text            =   "F_8"
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtMerchID 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ListBox CustList 
      Height          =   3180
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   7095
   End
   Begin VB.TextBox txtBTFID 
      Height          =   405
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdAuthTrans 
      Caption         =   "Authorize a Transaction"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearchByDealerID 
      Caption         =   "Seaech By Dealer ID"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearchByBTFID 
      Caption         =   "Search By BTF Number"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6480
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSearchByName 
      Caption         =   "Search by Company / Purchaser Name"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Label lblConnStatus 
      BackColor       =   &H80000011&
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "frmBTFCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private custArray() As String

Private Sub cmdAuthTrans_Click()
    If CustList.ListIndex <> -1 Then
        Dim index As Long
        Dim frmSale As frmBTFSale
        
        index = CustList.ListIndex
        Set frmSale = New frmBTFSale
        
        Load frmSale
        frmSale.setCustName (custArray(0, index))
        frmSale.setToken (custArray(1, index))
        frmSale.Show
        
    Else
        MsgBox "Please choose a token to use"
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSearchByName_Click()
    CustList.Clear
    If txtName.Text <> "" Then
        SearchByName txtName.Text
    Else
        MsgBox "Please Enter a purchaser name"
    End If
End Sub



Private Sub cmdSearchByBTFID_Click()
    Dim strRequest As String
    
    CustList.Clear
    
    If txtBTFID.Text = "" Then
        MsgBox "Please enter an id number"
    Else
        strRequest = modBTF.REQUESTPREFIX & "customers?bluetarp-cid=" & _
        txtBTFID.Text
        
        
        Label1.Caption = strRequest
        SendRequest strRequest, Inet1
    End If
End Sub


Private Sub cmdSearchByDealerID_Click()
    Dim strRequest As String
    CustList.Clear
    If txtMerchID.Text <> "" Then
    
        strRequest = modBTF.REQUESTPREFIX & "customers?merchant-cid=" & _
        txtMerchID.Text
    
        Label1.Caption = strRequest
        SendRequest strRequest, Inet1
    
    Else
        MsgBox "Must enter dealer's id for customer"
    End If
End Sub

Private Sub SearchByName(strName As String)
    Dim strRequest As String
    
    strRequest = modBTF.REQUESTPREFIX & "customers?q=" + strName
    
    Label1.Caption = strRequest
    SendRequest strRequest, Inet1
    
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
        End If
    End Select
End Sub

Private Sub parseResponse(strResponse As String)
    Dim doc As New MSXML.DOMDocument
    Dim success As Boolean
    Dim strTS As String
    
    success = doc.loadXML(strResponse)
    
    If success Then
        ReDim custArray(2, 10)
        Dim lngIndex As Long
        Dim nodeList As MSXML.IXMLDOMNodeList
        Set nodeList = doc.selectNodes("//bt:customers/*")
        If Not nodeList Is Nothing And nodeList.length > 0 Then
            Dim strCustName As String
            Dim strBTCID As String
            Dim strCred As String
        
            Dim nameNode As MSXML.IXMLDOMNode
            Dim BTNode As MSXML.IXMLDOMNode
            Dim credNode As MSXML.IXMLDOMNode
            Dim purchNodes As MSXML.IXMLDOMNodeList
            
            Dim node As MSXML.IXMLDOMNode
            
            lngIndex = 0
            
            For Each node In nodeList
                If node.nodeName = "bt:customer" Then
                    Set nameNode = node.selectSingleNode("bt:name")
                    Set BTNode = node.selectSingleNode("bt:number")
                    Set credNode = node.selectSingleNode("bt:available-credit")
                    strCustName = nameNode.Text
                    strBTCID = BTNode.Text
                    
                    If Not credNode Is Nothing Then
                       strCred = credNode.Text
                    Else
                        strCred = "    "
                    End If
                    Set purchNodes = node.selectNodes("bt:purchasers/*")
                    
                    Dim pNode As MSXML.IXMLDOMNode
                    Dim pName As MSXML.IXMLDOMNode
                    Dim pToken As MSXML.IXMLDOMNode
                    Dim strPName As String
                    Dim strToken As String
                    
                    For Each pNode In purchNodes
                         Set pName = pNode.selectSingleNode("bt:name")
                         Set pToken = pNode.selectSingleNode("bt:token")
                         
                         strPName = pName.Text
                         If Not pToken Is Nothing Then
                            Dim strEntry As String
                            strToken = pToken.Text
                            
                            strEntry = strCustName & "    " & strBTCID & _
                            "    " & strPName & "    " & strCred & "    " & _
                            strToken
                            CustList.AddItem (strEntry)
                            
                            If UBound(custArray, 2) = lngIndex Then
                                ReDim Preserve custArray(2, UBound(custArray, 2) + 10)
                            Else
                                custArray(0, lngIndex) = strCustName
                                custArray(1, lngIndex) = strToken
                            End If
                            lngIndex = lngIndex + 1
                         End If
                    Next pNode
                End If
                
            Next node
            
        Else
            Label1.Caption = "No Matches Found"
        End If
        
    Else
        lblConnStatus.Caption = "Unable to parse response"
    End If
        
        
End Sub

