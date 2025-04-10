Attribute VB_Name = "Graph"
Option Compare Database
Option Explicit

Private pClient As WebClient
Private pClientId As String
Private pTenantID As String
Private pWaitForLogin As Integer

Public Property Get Client() As WebClient
    If pClient Is Nothing Then
        Set pClient = New WebClient
        pClient.BaseUrl = "https://graph.microsoft.com/v1.0"
        pClientId = DLookup("ClientID", "AdminTable") 'Application (client) ID
        pTenantID = DLookup("TenantID", "AdminTable") 'Directory (Tenant) ID
        pWaitForLogin = DLookup("WaitForLogin", "AdminTable") 'Login wait period defaults to 60 seconds
        
        Dim Auth As New GraphAuthenticator
        Auth.Setup pClientId, pTenantID, pWaitForLogin
'        Auth.AddScope "offline_access"  'if using Refresh Token
        Auth.AddScope "mail.readwrite"
        Auth.AddScope "mail.send"
        Auth.AddScope "calendars.readwrite"
        Auth.AddScope "contacts.readwrite"
        Auth.AuthorizationUrl = "https://login.microsoftonline.com/" & pTenantID & "/oauth2/v2.0/authorize"
        Call Auth.Login
        
        Set pClient.Authenticator = Auth
    End If
    
    Set Client = pClient
End Property

Public Sub ClearAuthCodes()
    Dim Auth As New GraphAuthenticator
    Set Auth = Client.Authenticator
    Call Auth.ClearCodes
End Sub

Public Sub Logout()
    Dim Auth As New GraphAuthenticator
    Set Auth = Client.Authenticator
    Call Auth.Logout
End Sub


Public Function CreateDraftMessage(Subject As String, BodyType As String, BodyContent As String, toRecipients As String, ccRecipients As String, bccRecipients As String, AttachmentPath As String) As WebResponse
    Dim Request As New WebRequest
    Request.Resource = "/me/messages"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    
    Dim body As Dictionary
    Set body = New Dictionary
    body.Add "contentType", BodyType
    body.Add "content", BodyContent
    
    'Since toRecipients can be a list of email addresses
    Dim recipients As Collection
    Set recipients = New Collection
    FillEmailAddressCollection recipients, toRecipients
    
    Dim COLccRecipients As Collection
    If Trim(ccRecipients) <> "" Then
        Set COLccRecipients = New Collection
        FillEmailAddressCollection COLccRecipients, ccRecipients
    End If
    
    Dim COLbccRecipients As Collection
    If Trim(bccRecipients) <> "" Then
        Set COLbccRecipients = New Collection
        FillEmailAddressCollection COLbccRecipients, bccRecipients
    End If
    
    'Since attachments can be a list of attachments
    Dim attachments As Collection
    If Trim(AttachmentPath) <> "" Then
        Set attachments = New Collection
        Dim attachment As Dictionary
        Dim sFileName As String
        sFileName = Mid(AttachmentPath, InStrRev(AttachmentPath, "\") + 1)
        Set attachment = New Dictionary
        attachment.Add "@odata.type", "#microsoft.graph.fileAttachment"
        attachment.Add "name", sFileName
'        attachment.Add "contentType", "text/plain" 'Not mandatory so leave off for flexibility
        attachment.Add "contentBytes", ConvertFileToBase64(AttachmentPath)
        attachments.Add attachment, ""
    End If
    
    With Request
        .AddBodyParameter "subject", Subject
        .AddBodyParameter "body", body
        .AddBodyParameter "toRecipients", recipients
        If Trim(ccRecipients) <> "" Then .AddBodyParameter "ccRecipients", COLccRecipients
        If Trim(bccRecipients) <> "" Then .AddBodyParameter "bccRecipients", COLbccRecipients
        If Trim(AttachmentPath) <> "" Then .AddBodyParameter "attachments", attachments
    End With
    
    Dim sStatus As String
    sStatus = "Retry"
    While sStatus = "Retry"
        Set CreateDraftMessage = Client.Execute(Request)
        If CreateDraftMessage.StatusCode = WebStatusCode.Unauthorized And InStr(CreateDraftMessage.Content, "expired") Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend

End Function

Private Sub FillEmailAddressCollection(ByVal fillCollection As Collection, fillString As String)
    Dim sAddress As String
    Dim EmailAddress As Dictionary
    
    While InStr(fillString, ";") > 0
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        sAddress = Trim(Left(fillString, InStr(fillString, ";") - 1))
        fillString = Mid(fillString, InStr(fillString, ";") + 1)
        EmailAddress.Item("emailAddress").Add "address", sAddress
        fillCollection.Add EmailAddress
    Wend
    Set EmailAddress = New Dictionary
    EmailAddress.Add "emailAddress", New Dictionary
    EmailAddress.Item("emailAddress").Add "address", fillString
    fillCollection.Add EmailAddress
End Sub

Public Function GraphSendMail(Subject As String, BodyType As String, BodyContent As String, toRecipients As String, ccRecipients As String, bccRecipients As String, AttachmentPath As String) As WebResponse
    Dim Request As New WebRequest
    Request.Resource = "/me/sendMail"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    
    Dim message As Dictionary
    Set message = New Dictionary
    
    Dim body As Dictionary
    Set body = New Dictionary
    body.Add "contentType", BodyType
    body.Add "content", BodyContent
    
    'Since toRecipients can be a list of email addresses
    Dim recipients As Collection
    Set recipients = New Collection
    FillEmailAddressCollection recipients, toRecipients
    
    Dim COLccRecipients As Collection
    If Trim(ccRecipients) <> "" Then
        Set COLccRecipients = New Collection
        FillEmailAddressCollection COLccRecipients, ccRecipients
    End If
    
    Dim COLbccRecipients As Collection
    If Trim(bccRecipients) <> "" Then
        Set COLbccRecipients = New Collection
        FillEmailAddressCollection COLbccRecipients, bccRecipients
    End If
    
    'Since attachments can be a list of attachments
    Dim attachments As Collection
    If Trim(AttachmentPath) <> "" Then
        Set attachments = New Collection
        Dim attachment As Dictionary
        Dim sFileName As String
        sFileName = Mid(AttachmentPath, InStrRev(AttachmentPath, "\") + 1)
        Set attachment = New Dictionary
        attachment.Add "@odata.type", "#microsoft.graph.fileAttachment"
        attachment.Add "name", sFileName
'        attachment.Add "contentType", "text/plain" 'Not mandatory so leave off for flexibility
        attachment.Add "contentBytes", ConvertFileToBase64(AttachmentPath)
        attachments.Add attachment, ""
    End If
    
    With message
        .Add "subject", Subject
        .Add "body", body
        .Add "toRecipients", recipients
        If Trim(ccRecipients) <> "" Then .Add "ccRecipients", COLccRecipients
        If Trim(bccRecipients) <> "" Then .Add "bccRecipients", COLbccRecipients
        If Trim(AttachmentPath) <> "" Then .Add "attachments", attachments
    End With
    
    Request.AddBodyParameter "message", message
    Dim sStatus As String
    sStatus = "Retry"
    While sStatus = "Retry"
        Set GraphSendMail = Client.Execute(Request)
        If GraphSendMail.StatusCode = WebStatusCode.Unauthorized And InStr(GraphSendMail.Content, "expired") Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend

End Function

Public Function CreateGUID() As String
    Do While Len(CreateGUID) < 32
        If Len(CreateGUID) = 16 Then
            '17th character holds version information
            CreateGUID = CreateGUID & Hex$(8 + CInt(Rnd * 3))
        End If
        CreateGUID = CreateGUID & Hex$(CInt(Rnd * 15))
    Loop
    CreateGUID = Mid(CreateGUID, 1, 8) & "-" & Mid(CreateGUID, 9, 4) & "-" & Mid(CreateGUID, 13, 4) & "-" & Mid(CreateGUID, 17, 4) & "-" & Mid(CreateGUID, 21, 12)
End Function

Private Sub FillAttendeeCollection(ByVal fillCollection As Collection, fillStringReq As String, fillStringOpt As String)
    Dim sAddress As String
    Dim EmailAddress As Dictionary
    
    While InStr(fillStringReq, ";") > 0
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        sAddress = Trim(Left(fillStringReq, InStr(fillStringReq, ";") - 1))
        fillStringReq = Mid(fillStringReq, InStr(fillStringReq, ";") + 1)
        EmailAddress.Item("emailAddress").Add "address", sAddress
        EmailAddress.Add "type", "required"
        fillCollection.Add EmailAddress
    Wend
    If fillStringReq <> "" Then
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        EmailAddress.Item("emailAddress").Add "address", fillStringReq
        EmailAddress.Add "type", "required"
        fillCollection.Add EmailAddress
    End If
    While InStr(fillStringOpt, ";") > 0
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        sAddress = Trim(Left(fillStringOpt, InStr(fillStringOpt, ";") - 1))
        fillStringOpt = Mid(fillStringOpt, InStr(fillStringOpt, ";") + 1)
        EmailAddress.Item("emailAddress").Add "address", sAddress
        EmailAddress.Add "type", "optional"
        fillCollection.Add EmailAddress
    Wend
    If fillStringOpt <> "" Then
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        EmailAddress.Item("emailAddress").Add "address", fillStringOpt
        EmailAddress.Add "type", "optional"
        fillCollection.Add EmailAddress
    End If
End Sub

Public Function CreateEvent(Subject As String, BodyType As String, BodyContent As String, dStart As Date, tStart As Date, dEnd As Date, tEnd As Date, sLocation As String, sAttendees As String, sOptional As String) As WebResponse
    Dim Request As New WebRequest
    Request.Resource = "/me/events"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON

    Dim body As Dictionary
    Set body = New Dictionary
    body.Add "contentType", BodyType
    body.Add "content", BodyContent
    
    Dim start As Dictionary
    Set start = New Dictionary
    start.Add "dateTime", Format(dStart, "YYYY-MM-DD") + "T" + Format(tStart, "HH:MM:SS")
    start.Add "timeZone", Replace(CurrentTimeZone(), Chr(0), "")
    
    Dim enddic As Dictionary
    Set enddic = New Dictionary
    enddic.Add "dateTime", Format(dEnd, "YYYY-MM-DD") + "T" + Format(tEnd, "HH:MM:SS")
    enddic.Add "timeZone", Replace(CurrentTimeZone(), Chr(0), "")
    
    Dim location As Dictionary
    Set location = New Dictionary
    location.Add "displayName", sLocation
    
    'Since Attendees can be a list of email addresses
    Dim attendees As Collection
    Set attendees = New Collection
    FillAttendeeCollection attendees, sAttendees, sOptional
    
    With Request
        .AddBodyParameter "subject", Subject
        .AddBodyParameter "body", body
        .AddBodyParameter "start", start
        .AddBodyParameter "end", enddic
        If Trim(sLocation) <> "" Then .AddBodyParameter "location", location
        .AddBodyParameter "attendees", attendees
        .AddBodyParameter "allowNewTimeProposals", "true"
        .AddBodyParameter "transactionId", CreateGUID()
    End With
    
    Dim sStatus As String
    sStatus = "Retry"
    While sStatus = "Retry"
        Set CreateEvent = Client.Execute(Request)
        If CreateEvent.StatusCode = WebStatusCode.Unauthorized And InStr(CreateEvent.Content, "expired") Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
End Function

Public Function ListContacts(sFolder As String) As WebResponse
    Dim Request As New WebRequest
    
    If sFolder = "" Then
        Request.Resource = "/me/contacts"
    Else
        'Syntax for getting a specific folder uses the folder id so you have to get that first
        Request.Resource = "/me/contactfolders/" & GetFolderID(sFolder) & "/contacts"
    End If
    Request.Method = WebMethod.HttpGET
    Request.Format = WebFormat.JSON
    Request.AddQuerystringParam "Top", 1000
    
    Dim sStatus As String
    sStatus = "Retry"
    While sStatus = "Retry"
        Set ListContacts = Client.Execute(Request)
        If ListContacts.StatusCode = WebStatusCode.Unauthorized And InStr(ListContacts.Content, "expired") Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
End Function

Public Function GetFolderID(sFolder As String) As String
    Dim Request As New WebRequest
    Dim Response As New WebResponse
    
    GetFolderID = "Retry"
    While GetFolderID = "Retry"
        Request.Resource = "/me/contactFolders"
        Set Response = Client.Execute(Request)
        If Response.StatusCode = WebStatusCode.OK Then
            Dim FolderInfo As Dictionary
            For Each FolderInfo In Response.Data("value")
                If FolderInfo("displayName") = sFolder Then
                    GetFolderID = FolderInfo("id")
                End If
            Next FolderInfo
        Else
            If Response.StatusCode = WebStatusCode.Unauthorized And InStr(Response.Content, "expired") Then
                ClearAuthCodes
                GetFolderID = "Retry"
            Else
                MsgBox "Error " & Response.StatusCode & ": " & Response.Content
                GetFolderID = "Error"
            End If
        End If
    Wend
End Function

Public Function CreateContact(sFolder As String, givenName As String, surname As String, fileAs As String, jobTitle As String, companyName As String, sBusinessPhones As String, sEmailAddresses As String) As WebResponse
    Dim Request As New WebRequest
    If sFolder = "" Then
        Request.Resource = "/me/contacts"
    Else
        'Syntax for getting a specific folder uses the folder id so you have to get that first
        Request.Resource = "/me/contactfolders/" & GetFolderID(sFolder) & "/contacts"
    End If
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    
    'Since emailAddresses can be a list of email addresses
    Dim emailAddresses As Collection
    Set emailAddresses = New Collection
    Dim sAddress As String
    Dim EmailAddress As Dictionary
    
    While InStr(sEmailAddresses, ";") > 0
        Set EmailAddress = New Dictionary
        sAddress = Trim(Left(sEmailAddresses, InStr(sEmailAddresses, ";") - 1))
        sEmailAddresses = Mid(sEmailAddresses, InStr(sEmailAddresses, ";") + 1)
        EmailAddress.Add "address", sAddress
        emailAddresses.Add EmailAddress
    Wend
    Set EmailAddress = New Dictionary
    EmailAddress.Add "address", sEmailAddresses
    emailAddresses.Add EmailAddress
    
    'Since businessPhones can be a list of phone numbers
    Dim businessPhones() As String
    Dim sPhone As String
    Dim iPhoneCount As Integer
    If Trim(sBusinessPhones) <> "" Then
        While InStr(sBusinessPhones, ";") > 0
            ReDim businessPhones(iPhoneCount)
            sPhone = Trim(Left(sBusinessPhones, InStr(sBusinessPhones, ";") - 1))
            sBusinessPhones = Mid(sBusinessPhones, InStr(sBusinessPhones, ";") + 1)
            businessPhones(iPhoneCount) = sPhone
            iPhoneCount = iPhoneCount + 1
        Wend
        ReDim businessPhones(iPhoneCount)
        businessPhones(iPhoneCount) = sBusinessPhones
        iPhoneCount = iPhoneCount + 1
    End If
    
    With Request
        .AddBodyParameter "givenName", givenName
        .AddBodyParameter "surname", surname
        .AddBodyParameter "fileAs", fileAs
        .AddBodyParameter "jobTitle", jobTitle
        .AddBodyParameter "companyName", companyName
        .AddBodyParameter "emailAddresses", emailAddresses
        If iPhoneCount > 0 Then .AddBodyParameter "businessPhones", businessPhones
    End With
    
    Dim sStatus As String
    sStatus = "Retry"
    While sStatus = "Retry"
        Set CreateContact = Client.Execute(Request)
        If CreateContact.StatusCode = WebStatusCode.Unauthorized And InStr(CreateContact.Content, "expired") Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
End Function

