VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Soap Client"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab 
      Height          =   5025
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   8864
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Client Test"
      TabPicture(0)   =   "frmMain.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblServerAddy"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTexttoSend"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTextResponse"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtServerAddy"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtToSend"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtReturned"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSend"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdClose"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Soap Request"
      TabPicture(1)   =   "frmMain.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtSoapRequest"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Soap Response"
      TabPicture(2)   =   "frmMain.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtSoapResponse"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtSoapResponse 
         Height          =   4365
         Left            =   -74820
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   480
         Width           =   5955
      End
      Begin VB.TextBox txtSoapRequest 
         Height          =   4365
         Left            =   -74820
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   480
         Width           =   5955
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   405
         Left            =   4890
         TabIndex        =   8
         Top             =   4320
         Width           =   1155
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   405
         Left            =   3630
         TabIndex        =   7
         Top             =   4320
         Width           =   1155
      End
      Begin VB.TextBox txtReturned 
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   2640
         Width           =   3945
      End
      Begin VB.TextBox txtToSend 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1830
         Width           =   3945
      End
      Begin VB.TextBox txtServerAddy 
         Height          =   285
         Left            =   330
         TabIndex        =   1
         Text            =   "http://localhost/soapserver/soaprequest.asp"
         Top             =   1020
         Width           =   3945
      End
      Begin VB.Label lblTextResponse 
         Caption         =   "Text Returned"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2430
         Width           =   1215
      End
      Begin VB.Label lblTexttoSend 
         Caption         =   "Text to Send"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1620
         Width           =   1545
      End
      Begin VB.Label lblServerAddy 
         Caption         =   "Server Address"
         Height          =   255
         Left            =   330
         TabIndex        =   2
         Top             =   810
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

End

End Sub

Private Sub cmdSend_Click()

Dim objxmlhttp     As New MSXML.XMLHTTPRequest
Dim objxml         As New MSXML.DOMDocument
Dim objroot        As IXMLDOMElement
Dim objxmlresponse As IXMLDOMNode
Dim strSOAPToSend  As String

txtReturned.Text = ""
txtSoapRequest.Text = ""
txtSoapResponse = ""

SSTab.TabEnabled(1) = False
SSTab.TabEnabled(2) = False

objxmlhttp.open "POST", txtServerAddy, False
objxmlhttp.setRequestHeader "Man", POST & " " & txtServerAddy & " HTTP/1.1"
objxmlhttp.setRequestHeader "MessageType", "CALL"
objxmlhttp.setRequestHeader "ContentType", "text/xml"
strSOAPToSend = BuildSOAP
objxmlhttp.send strSOAPToSend

If objxmlhttp.Status = 200 Then
    Set objxml = objxmlhttp.responseXML
    Set objroot = objxml.documentElement
    Set objxmlresponse = objroot.selectSingleNode(".//response")
    
    txtReturned.Text = objxmlresponse.Text
    txtSoapRequest.Text = strSOAPToSend
    txtSoapResponse = objxmlhttp.responseText
    SSTab.TabEnabled(1) = True
    SSTab.TabEnabled(2) = True
Else
    txtReturned.Text = objxmlhttp.statusText
End If

Set objxmlhttp = Nothing
Set objxml = Nothing

End Sub

Private Sub Form_Load()

SSTab.TabEnabled(1) = False
SSTab.TabEnabled(2) = False

End Sub

Private Sub SSTab1_DblClick()

End Sub


Public Function BuildSOAP() As String

Dim strSOAP As String

strSOAP = "<SOAP:Envelope  " & _
          "xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
strSOAP = strSOAP & "<SOAP:Body>"
strSOAP = strSOAP & "<Name>SoapDest.cStringTest</Name>"
strSOAP = strSOAP & "<Proc>StringReverse</Proc>"
strSOAP = strSOAP & "<Params>" & txtToSend & "</Params>"
strSOAP = strSOAP & "</SOAP:Body>"
strSOAP = strSOAP & "</SOAP:Envelope>"

BuildSOAP = strSOAP

End Function
