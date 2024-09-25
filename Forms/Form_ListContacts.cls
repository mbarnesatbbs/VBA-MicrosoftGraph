VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ListContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    CurrentDb.Execute "DELETE FROM TempContactsTable"
    
    'parameter is optional, if filled in use the name of a folder,
    'for example my contacts are in BusinessContacts so I use that
    ListOutlookContacts Graph.ListContacts("")
    Me.lstContacts.Requery
End Sub

Private Sub ListOutlookContacts(Response As WebResponse)
    If Response.StatusCode = WebStatusCode.OK Then
        Dim ContactInfo As Dictionary
        Dim EmailColl As Collection
        Dim EmailAddress As Dictionary
        Dim rs As Recordset
        
        If IsEmpty(Response.Data("value")) Then
            MsgBox "No contacts for this folder in this tenant."
        Else
            Set rs = CurrentDb.OpenRecordset("TempContactsTable")
            For Each ContactInfo In Response.Data("value")
                Set EmailColl = ContactInfo("emailAddresses")
                For Each EmailAddress In EmailColl
                    If EmailAddress("address") <> "" Then
                        rs.AddNew
                        rs.Fields("id").Value = ContactInfo("id")
                        rs.Fields("createdDateTime").Value = ContactInfo("createdDateTime")
                        rs.Fields("lastModifiedDateTime").Value = ContactInfo("lastModifiedDateTime")
                        rs.Fields("displayName").Value = ContactInfo("displayName")
                        rs.Fields("givenName").Value = ContactInfo("givenName")
                        rs.Fields("surname").Value = ContactInfo("surname")
                        rs.Fields("emailAddress").Value = EmailAddress("address")
                        rs.Update
                    End If
                Next EmailAddress
            Next ContactInfo
            rs.Close
        End If
        Set rs = Nothing
    Else
        If Response.StatusCode = WebStatusCode.Unauthorized And InStr(Response.Content, "expired") Then ClearAuthCodes
        MsgBox "Error " & Response.StatusCode & ": " & Response.Content
    End If
End Sub
