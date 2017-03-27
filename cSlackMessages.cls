VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSlackMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mChannel As String
Private mWebHookUrl As String
Private mDebutMessage As String
Private mFinMessage As String
Private mWebhookSurname As String
Private mDirectMessageUsername As String
Private mHypertextLink As String
Private mIconUrl As String
Private mIconEmoji As String


Property Get strChannel() As String
strChannel = mChannel
End Property

Property Let strChannel(strChannel As String)

If Left(strChannel, 1) <> "#" Then
mChannel = "#" & strChannel
Else
mChannel = strChannel
End If


End Property

Property Get strDirectMessageUsername() As String
strDirectMessageUsername = mDirectMessageUsername
End Property

Property Let strDirectMessageUsername(strDirectMessageUsername As String)

If Left(strDirectMessageUsername, 1) <> "@" Then
mDirectMessageUsername = "@" & strDirectMessageUsername
Else
mDirectMessageUsername = strDirectMessageUsername
End If
If strDirectMessageUsername <> "" Then mChannel = mDirectMessageUsername

End Property

Property Get strwebHookUrl() As String
strwebHookUrl = mWebHookUrl
End Property
Property Let strwebHookUrl(strwebHookUrl As String)
mWebHookUrl = strwebHookUrl
End Property

Property Get strwebhookSurname() As String
strwebhookSurname = mWebhookSurname
End Property
Property Let strwebhookSurname(strwebhookSurname As String)
mWebhookSurname = strwebhookSurname
End Property
Property Get strdebutMessage() As String
strdebutMessage = mDebutMessage
End Property
Property Let strdebutMessage(strdebutMessage As String)
mDebutMessage = strdebutMessage
End Property
Property Get strfinMessage() As String
strfinMessage = mFinMessage
End Property
Property Let strfinMessage(strfinMessage As String)
mFinMessage = strfinMessage
End Property
Property Get strHypertextLink() As String
strHypertextLink = mHypertextLink
End Property
Property Let strHypertextLink(strHypertextLink As String)
'Web site URL  like "http://www.google.com"
mHypertextLink = "<" & strHypertextLink & ">"
End Property
Property Get message() As String
message = strdebutMessage & "\n" & strHypertextLink & "\n" & strfinMessage
End Property

Property Get striconUrl() As String
striconUrl = mIconUrl
End Property
Property Let striconUrl(striconUrl As String)
mIconUrl = striconUrl
End Property

Property Get striconEmoji() As String
striconEmoji = mIconEmoji
End Property

Property Let striconEmoji(striconEmoji As String)
mIconEmoji = striconEmoji
End Property
Property Get PostSlack() As String

Dim httpReq As Object
  Set httpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
   With httpReq
   .Option(4) = 13056 'WinHttpRequestOption_SslErrorIgnoreFlags
   .Open "POST", strwebHookUrl, False
   .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
   Debug.Print jSon
   .Send (jSon)
   PostSlack = .responseText
   End With
   
   Set httpReq = Nothing
End Property

Property Get jSon() As String
jSon = "{"
  jSon = jSon & " ""channel"": """ & strChannel & ""","
  If strwebhookSurname <> "" Then jSon = jSon & " ""username"": """ & strwebhookSurname & ""","
  jSon = jSon & " ""text"": """ & message & """"
    jSon = jSon & "}"
End Property
