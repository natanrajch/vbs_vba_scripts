
'This should go in the "ThisOutlookSession" of Microsoft Outlook
'It runs custom subs when some Outlook events are fired. 
'Mainly built for scheduled Appointments/CITAS with a custom label/ETIQUETA (Category).
Private Sub Application_Reminder(ByVal item As Object)
  Dim objMsg As MailItem
  Set objMsg = Application.CreateItem(olMailItem)
  Dim sHoy As String
 
sHoy = FormatDateTime(Now(), vbLongDate)
 
If item.MessageClass <> "IPM.Appointment" Then
'    If item.MessageClass = "IPM.Note" And _
'        TypeName(item) = "Mailitem" Then
'            
'            Call GuardaAdjuntosCoordinaciones(item)
'    End If
'Else
  Exit Sub
End If

If item.Categories = "ETIQUETA 1" Then
  Call ReporteSemanal
End If

If item.Categories = "ETIQUETA 2" Then
  Call another_possible_sub
End If

If item.Categories = "ETIQUETA 3" Then
  Call another_possible_sub
End If

If item.Categories <> "Mail Automatizado" Then
  Exit Sub
End If


End Sub

'Option Explicit
' New Email arriving Event:


Private Sub Application_NewMail()

 Dim i As Integer
 Dim CuentaMails As Integer
 Dim myBandeja As Outlook.Folder
 
  
 
 Set myBandeja = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
 CuentaMails = myBandeja.Items.Count
 
For i = 0 To 15
 
 If TypeName(myBandeja.Items(CuentaMails - i)) = "MailItem" Then
    If myBandeja.Items(CuentaMails - i).Sender = "SENDER NAME / NOT ADDRESS" Then
    
        Call GuardaAdjuntos(myBandeja.Items(CuentaMails - i))
    End If
 End If
Next

 Set myBandeja = Nothing
 
 End Sub
