VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSoapHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private strSOAPResponse As String

Public Function ProcessSOAP(hRequest As Object) As Boolean

'XML Objects
Dim objXML          As New MSXML.DOMDocument
Dim objRoot         As MSXML.IXMLDOMElement
Dim objName         As MSXML.IXMLDOMNode
Dim objMethod       As MSXML.IXMLDOMNode
Dim objParamList    As MSXML.IXMLDOMNodeList
Dim Node            As MSXML.IXMLDOMNode

'General Vars
Dim retval          As Boolean
Dim strResponse     As String
Dim strObjectName   As String
Dim strObjectMethod As String
Dim Params()        As Variant
Dim iWorked         As Boolean

'Redim Params Array
ReDim Params(0)

'Load the XML
retval = objXML.Load(hRequest)

If retval = True Then
    'Process the Soap
    'Load XML Nodes
    Set objRoot = objXML.documentElement
    Set objName = objRoot.selectSingleNode(".//Name")
    Set objMethod = objRoot.selectSingleNode(".//Proc")
    Set objParamList = objRoot.selectNodes(".//Params")
        
    'Assign Node Values to Variables
    strObjectName = objName.Text
    strObjectMethod = objMethod.Text
    Open "c:\test.txt" For Append As #1
    Print #1, strObjectName & " -- " & strObjectMethod
    Close #1
    
    For Each Node In objParamList
        ReDim Preserve Params(UBound(Params) + 1)
        Params(UBound(Params) - 1) = Node.Text
    Next
    
    iWorked = MethodToExec(strObjectName, strObjectMethod, Params())
    
    If iWorked = True Then
        ProcessSOAP = True
        GoTo TidyUp
    Else
        ProcessSOAP = False
        GoTo TidyUp
    End If
Else
    ProcessSOAP = False
    GoTo TidyUp
End If

TidyUp:
    Set objParamList = Nothing
    Set objMethod = Nothing
    Set objName = Nothing
    Set objRoot = Nothing
    Set objXML = Nothing
    
    Exit Function

End Function

Private Function MethodToExec(objName As String, Method As String, Params() As Variant) As Boolean

Dim strObjName      As String
Dim strObjMethod    As String
Dim vParamList      As Variant
Dim strResponse     As String
'Dim strSOAPResponse As String


'Parse Params
For x = 0 To UBound(Params) - 1
    If x = 0 Then
        vParamList = Params(x)
    Else
        vParamList = vParamList & ", " & Params(x)
    End If
Next

Set objToCall = CreateObject(objName)

If IsObject(objToCall) Then
    
    strResponse = CallByName(objToCall, Method, VbMethod, vParamList)
    
    strSOAPResponse = strSOAPResponse & "<SOAP:Envelope " & _
                      "xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
    strSOAPResponse = strSOAPResponse & "<SOAP:Body>"
    strSOAPResponse = strSOAPResponse & "<response>" & strResponse & "</response>"
    strSOAPResponse = strSOAPResponse & "</SOAP:Body>"
    strSOAPResponse = strSOAPResponse & "</SOAP:Envelope>"
    
    MethodToExec = True
Else
    MethodToExec = False
End If

End Function

Private Sub Class_Terminate()


End Sub



Public Property Get SOAPResponse() As String

SOAPResponse = strSOAPResponse

End Property

