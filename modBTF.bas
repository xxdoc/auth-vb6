Attribute VB_Name = "modBTF"
Public Const HOSTNAME As String = "integration.bluetarp.com"
Public Const MERCHID As String = "7"
Public Const CLIENTKEY As String = "wLn1zexQS3dzIGx3bG2U2M"
Public Const REQUESTPREFIX As String = HOSTNAME & "/auth/v1/" & _
MERCHID & "/"
Public Const btns = "{http://api.bluetarp.com/ns/1.0}"

Public Enum AuthRequestType
    sale
    credit
    DepositHold
    DepositCollect
    Void
End Enum


Public Function CreateXMLRequest(ByRef btTrans As Object) As String
    Dim reqDoc As MSXML.DOMDocument
    Dim procInst As MSXML.IXMLDOMProcessingInstruction
    
    'Create document and add root
    Set reqDoc = New MSXML.DOMDocument
    Set procInst = reqDoc.createProcessingInstruction("xml", "version='1.0' encoding ='utf-8'")
    reqDoc.appendChild procInst
    
    Dim root As MSXML.IXMLDOMElement
    Set root = reqDoc.createNode(NODE_ELEMENT, "bt:bluetarp-authorization", "http://api.bluetarp.com/ns/1.0")
    reqDoc.appendChild root
    
    'Create and add properties to root
    Dim xsins As IXMLDOMAttribute
    Dim xsischema As IXMLDOMAttribute
    Dim btns As IXMLDOMAttribute
    
    'Set btns = reqDoc.createAttribute("xmlns:bt")
    'root.appendChild btns
    'btns.
    
    'Create auth-request elem and add to root
    Dim authreq As MSXML.IXMLDOMElement
    Set authreq = reqDoc.createElement("bt:authorization-request")
    root.appendChild authreq
    
    'Add and set merchant id
    Dim merchnum As MSXML.IXMLDOMElement
    Set merchnum = reqDoc.createElement("bt:merchant-number")
    authreq.appendChild merchnum
    merchnum.Text = modBTF.MERCHID
    
    'Add and set client id
    Dim clientid As MSXML.IXMLDOMElement
    Set clientid = reqDoc.createElement("bt:client-id")
    authreq.appendChild clientid
    clientid.Text = btTrans.custName
    
    'Add and set transaction id
    Dim transid As MSXML.IXMLDOMElement
    Set transid = reqDoc.createElement("bt:transaction-id")
    authreq.appendChild transid
    transid.Text = btTrans.transid
    
    'Add and set purchaser and token
    Dim purchaser As MSXML.IXMLDOMElement
    Dim token As MSXML.IXMLDOMElement
    Set purchaser = reqDoc.createElement("bt:purchaser-with-token")
    Set token = reqDoc.createElement("bt:token")
    authreq.appendChild purchaser
    purchaser.appendChild token
    token.Text = btTrans.token
    
    'Add transaction element
    Dim trans As MSXML.IXMLDOMElement
    Dim oldInvoice As MSXML.IXMLDOMElement
    Dim authseq As MSXML.IXMLDOMElement
    
    Select Case btTrans.tType
        Case AuthRequestType.sale
            Set trans = reqDoc.createElement("bt:sale")
        Case AuthRequestType.credit
            Set trans = reqDoc.createElement("bt:credit")
        Case AuthRequestType.DepositCollect
            Set trans = reqDoc.createElement("bt:deposit-collect")
        Case AuthRequestType.DepositHold
            Set trans = reqDoc.createElement("bt:deposit-hold")
        Case AuthRequestType.Void
            Set trans = reqDoc.createElement("bt:void")
            Set authseq = reqDoc.createElement("bt:auth-seq")
            trans.appendChild authseq
            authseq.Text = btTrans.authseq
    End Select
    
    authreq.appendChild trans
    
    If btTrans.tType <> Void Then
        Dim amount As MSXML.IXMLDOMElement
        Set amount = reqDoc.createElement("bt:amount")
        trans.appendChild amount
        amount.Text = btTrans.amount
        
        If btTrans.tType = DepositCollect Then
            Set authseq = reqDoc.createElement("bt:auth-seq")
            trans.appendChild authseq
            authseq.Text = btTrans.authseq
        End If
        
        Dim jobid As MSXML.IXMLDOMElement
        Set jobid = reqDoc.createElement("bt:job-code")
        trans.appendChild jobid
        jobid.Text = btTrans.jobid
        
        Dim invoice As MSXML.IXMLDOMElement
        Set invoice = reqDoc.createElement("bt:invoice")
        trans.appendChild invoice
        invoice.Text = btTrans.invoice
        
        If btTrans.tType = credit Then
            Dim oldInv As MSXML.IXMLDOMElement
            Set oldInv = reqDoc.createElement("bt:original-invoice")
            trans.appendChild oldInv
            oldInv.Text = btTrans.oldInvoice
        End If
        
    End If
    
    CreateXMLRequest = reqDoc.xml
End Function

Public Sub SendRequest(strRequestURL As String, Inet1 As Inet)
    Dim strRequestHeader As String
    Dim strChunk As String
       
    strResponse = ""
    
    strRequestHeader = "Authorization: Bearer " & CLIENTKEY
    
    Inet1.Protocol = icHTTPS
    Inet1.Execute strRequestURL, "GET", "", strRequestHeader
    
End Sub

Public Sub SendAuthPost(strRequestURL As String, _
    strRequestXML As String, Inet1 As Inet)
    Dim strRequestHeader As String
    Dim strChunk As String
       
    strResponse = ""
    
    strRequestHeader = "Authorization: Bearer " & _
    CLIENTKEY & vbCrLf & "Content-Type: text/xml" & vbCrLf & _
    "Content-Encoding: UTF-8"
    
    Inet1.Protocol = icHTTPS
    Inet1.Execute strRequestURL, "POST", strRequestXML, strRequestHeader
    
End Sub

Public Function ParseTransList(strResponse As String) As String()
    Dim doc As New MSXML.DOMDocument
    Dim success As Boolean
    Dim transArray() As String
    
    success = doc.loadXML(strResponse)
    
    If success Then
        Dim nodeList As MSXML.IXMLDOMNodeList
        Dim lngCapacity As Long
        
        ReDim transArray(4, 10)
        lngIndex = 0
        Set nodeList = doc.selectNodes("//bt:transactions/*")
        
        If Not nodeList Is Nothing And nodeList.length <> 0 Then
            Dim node As MSXML.IXMLDOMNode
            Dim strCustName As String
            Dim strAuthSeq As String
            Dim strAuthToken As String
            Dim strAmount As String
            
            For Each node In nodeList
                
                strAuthSeq = node.selectSingleNode("bt:auth-seq").Text
                strAmount = node.selectSingleNode("bt:amount").Text
                strCustName = ""
                strAuthToken = ""
                
                Dim custNode As MSXML.IXMLDOMNode
                Set custNode = node.selectSingleNode("bt:customer")
                If Not custNode Is Nothing Then
                    strCustName = custNode.selectSingleNode("bt:name").Text
                    Dim purchNodes As MSXML.IXMLDOMNodeList
                    Set purchNodes = custNode.selectNodes("bt:purchasers/*")
                    
                    If Not purchNodes Is Nothing And purchNodes.length > 0 Then
                        Dim pNode As MSXML.IXMLDOMNode
                        Set pNode = purchNodes.nextNode
                        If Not pNode Is Nothing Then
                            strAuthToken = pNode.selectSingleNode("bt:token").Text
                        End If
                    End If
                End If
            
                transArray(0, lngIndex) = strAuthSeq
                transArray(1, lngIndex) = strAmount
                transArray(2, lngIndex) = strAuthToken
                transArray(3, lngIndex) = strCustName
                
                lngIndex = lngIndex + 1
                If lngIndex = UBound(transArray, 2) Then
                    ReDim Preserve transArray(4, UBound(transArray, 2) + 10)
                End If
               
            Next node
            ParseTransList = transArray
        Else
            Dim strNone() As String
            strNone(0) = "No Matches Found"
            ParseTransList = strNone
        End If
    Else
        Dim strErrorMsg() As String
        
        strErrorMsg(0) = "Failed to parse XML response"
        
        ParseTransList = strErrorMsg
        
    End If
    
    
End Function

Public Function ParseAuthResponse(strResponse As String) As String
    Dim doc As New MSXML.DOMDocument
    Dim success As Boolean
    
    success = doc.loadXML(strResponse)
    
    If success Then
        Dim codeNode As MSXML.IXMLDOMNode
        Dim msgNode As MSXML.IXMLDOMNode
        Dim authNode As MSXML.IXMLDOMNode
        Dim strCode As String
        Dim strMsg As String
        
        Set codeNode = doc.selectSingleNode("//bt:code")
        Set msgNode = doc.selectSingleNode("//bt:message")
        
        strCode = codeNode.Text
        strMsg = msgNode.Text
            
        If strCode = "00" Then
            Dim strAuthSeq As String
            Set authNode = doc.selectSingleNode("//bt:auth-seq")
            strAuthSeq = authNode.Text
            ParseAuthResponse = "Transaction was " & _
            strMsg & ":  " & strAuthSeq
        Else
            ParseAuthResponse = "Transaction was " & strMsg
        End If
        
    Else
        ParseAuthResponse = "Failed to parse XML response:" & _
        vbCrLf & "Please contact BlueTarp"
        
    End If
End Function
