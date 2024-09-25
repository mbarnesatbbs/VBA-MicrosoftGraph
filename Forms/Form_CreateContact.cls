VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CreateContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCreateContact_Click()
    If IsNull(fileAs) Or IsNull(sEmailAddresses) Then
        MsgBox "File As and Email are mandatory"
        Exit Sub
    End If
    'First parameter should be either an empty string "" or use the name of a folder
    CreateGraphContact Graph.CreateContact("BusinessContacts", Nz(Me.givenName, ""), Nz(Me.surname, ""), Me.fileAs, Nz(Me.jobTitle, ""), Nz(Me.companyName, ""), Nz(sBusinessPhones, ""), Me.sEmailAddresses)
End Sub

Private Sub CreateGraphContact(Response As WebResponse)
    If Response.StatusCode = WebStatusCode.Created Then
        MsgBox "Contact created."
    Else
        If Response.StatusCode = WebStatusCode.Unauthorized And InStr(Response.Content, "expired") Then ClearAuthCodes
        MsgBox "Error " & Response.StatusCode & ": " & Response.Content
    End If
End Sub

Private Sub surname_AfterUpdate()
    If IsNull(Me.fileAs) Then Me.fileAs = Me.surname & ", " & Me.givenName
End Sub
