
'Use these in a Microsoft Outlook Module 
Public Sub weekly_report()
  Dim objMsg As MailItem
  Dim dViernesInforme As Double
  Set objMsg = Application.CreateItem(olMailItem)
  
  dViernesInforme = Date - Weekday(Date, vbFriday) + 1

'Crea un mail, adjunta con valor 0 (invisible) los gráficos y los embedde en el html de manera ordenada.
'Creates an email, attaches pictures directly into the body (they wont be shown as attachments) 
          With objMsg
               .Subject = "YOUR SUBJECT - WEEK " & Format((dViernesInforme - 7), "dd-mm") & " al " & Format((dViernesInforme - 1), "dd-mm")
               .To = "recipient1@sth.com.ar;recipient2@sth.com.ar"
               '.CC = "recipient3@sth.com.ar"
               .HTMLBody = "<HTML><BODY><font size=""3"" face=""Calibri"">" & _
         "Greetings" & "<br>" & _
          "Here is your weekly update." & "<br>" & _
          "</font></HTML></BODY>"
               .Attachments.Add "PATH & image you want to add to html email", olByValue, 0
               .Attachments.Add "another one if you please", olByValue, 0
                                             
               '.Importance = olImportanceHigh
'               .ReadReceiptRequested = False
               .HTMLBody = .HTMLBody _
               'Now lets add a fancy html table so the email looks nicer
                & "<table style=""width:20%""><tr><th><img src='cid:images_file_name.png'" & "width='500' height='345'></th><th><img src='cid:anotherFileName.png'" & "width='500' height='345'></th></tr></table>"
               '.Display
               .Send
            End With
  Set objMsg = Nothing
End Sub

Sub GuardaAdjuntos(ByVal email As Object)
On Error GoTo ErrorHandler
Dim Msg As Outlook.MailItem
Dim olAdj As Object

If email.Sender = "SENDER NAME, NOT ADDRESS" _
    And email.Attachments.Count > 0 Then
        For Each olAdj In email.Attachments
            If olAdj.FileName Like "*.xls" Or olAdj.FileName Like "*.xlsx" Then
                olAdj.SaveAsFile "A-PATH" & olAdj.FileName
            End If
        Next
End If

Exit Sub

ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description

End Sub

'Original function written by Diane Poremsky: http://www.slipstick.com/developer/send-email-outlook-reminders-fires/