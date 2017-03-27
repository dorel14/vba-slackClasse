Attribute VB_Name = "Module2"
Sub test()
Dim mslack As New cSlackMessages

With mslack
    .webHookUrl = "https://hooks.slack.com/services/xxxxxxxxxxxxxxxxxxxxxx"
    .channel = "general"
    .debutMessage = "debut de message"
    .HypertextLink = "http://www.google.com"
    .finMessage = "fin  de message"
    .DirectMessageUsername = "@specificUser"
    .webhookSurname = "test-vba-excel"
    
End With

With mslack
Debug.Print .message
Debug.Print .PostSlack
End With

End Sub
