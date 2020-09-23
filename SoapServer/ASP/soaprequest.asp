<%@Language = "VBScript"%>
<%
Response.Buffer = True
Response.ContentType = "text/xml"

Dim objSoapServer
Dim retVal

Set objSoapServer = CreateObject("SoapServer.CSoapHandler")

'Response.Write isObject(objSoapServer)

'objSoapServer.Init Request, Response

retVal = objSoapServer.ProcessSOAP(Request)

If retVal = true then
	Response.Write objSoapServer.SoapResponse
Else
	Response.Write "<xml><response>Error Occurred</response></xml>"
End If

Set objSoapServer = Nothing

%>
