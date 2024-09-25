VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Written by Maria Barnes, Barnes Business Solutions, Inc.
'https://BarnesBusinessSolutions.com
' @class GraphAuthenticator
' @author mbarnes@BarnesBusinessSolutions.com

'Started from Google Authenticator in
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Interface for creating authenticators for rest client
'
' @class IWebAuthenticator
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)

Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Public AuthorizationUrl As String
Private Const auth_RedirectUrl As String = "https://login.microsoftonline.com/common/oauth2/nativeclient"

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ClientId As String
Public TenantID As String
Public ApiKey As String
Public AuthorizationCode As String
Public Token As String
Public Scopes As Variant
Public WaitForLogin As Integer

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup
'
' @param {String} ClientId
' @param {String} TenantID
''
Public Sub Setup(ClientId As String, TenantID As String, WaitForLogin As Integer)
    Me.ClientId = ClientId
    Me.TenantID = TenantID
    Me.WaitForLogin = WaitForLogin
End Sub

''
' ClearCodes
'
''
Public Sub ClearCodes()
    Me.AuthorizationCode = ""
    Me.Token = ""
End Sub


''
' Login to Microsoft Graph
''
' If you want to use IE from Windows or Mac, then uncomment this Login Sub and comment out the one below it
'     which is the default designed to be used from Windows with the New Edge Browser control
'Public Sub Login()
'    On Error GoTo auth_ErrorHandling
'
'    ' No need to login if API key, authorization code, or token have been set
'    If Me.ApiKey <> "" Or Me.AuthorizationCode <> "" Or Me.Token <> "" Then
'        Exit Sub
'    End If
'
'    Dim auth_Completed As Boolean
'    auth_Completed = True
'
'#If Mac Then
'
'    ' Mac login opens dialog and then user copy-paste's authorization code into InputBox
'    Dim auth_Result As ShellResult
'    Dim auth_Response As String
'
'    auth_Result = WebHelpers.ExecuteInShell("open " & WebHelpers.PrepareTextForShell(GetLoginUrl()))
'
'    If auth_Result.ExitCode <> 0 Then
'        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Unable to open browser"
'    End If
'
'    auth_Response = VBA.InputBox("Opening Microsoft Graph Login..." & vbNewLine & vbNewLine & _
'        "After you've logged in, copy the code from the browser and paste it here to authorize this application", _
'        Title:="Logging in...")
'
'    If auth_Response = "" Then
'        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login was cancelled"
'    End If
'
'    ' Success!
'    Me.AuthorizationCode = auth_Response
'
'#Else
'
'    ' Windows login uses IE to automate retrieving authorization code for user
'    Dim auth_ErrorDescription As String
'    On Error GoTo auth_Cleanup
'
'    auth_Completed = False
'
'    Dim auth_IE As Object
'    Set auth_IE = CreateObject("InternetExplorer.Application")
'    auth_IE.Silent = True
'    auth_IE.AddressBar = False
'    auth_IE.Navigate GetLoginUrl()
'    auth_IE.Visible = True
'
'    ' Wait for login to complete
'    Do While Not IEauth_LoginIsComplete(auth_IE)
'        DoEvents
'    Loop
'    auth_Completed = True
'
'    If IEauth_LoginIsDenied(auth_IE) Then
'        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login failed or was denied"
'    ElseIf IEauth_LoginIsError(auth_IE) Then
'        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login error: " & IEauth_LoginExtractError(auth_IE)
'    End If
'
'    ' Success!
'    Me.AuthorizationCode = auth_LoginExtractCode(auth_IE)
'
'auth_Cleanup:
'
'    If Not auth_IE Is Nothing Then: auth_IE.Quit
'    Set auth_IE = Nothing
'
'#End If
'
'    If Err.Number = 0 And auth_Completed Then
'        WebHelpers.LogDebug "Login succeeded: " & Me.AuthorizationCode, "GraphAuthenticator.Login"
'        Exit Sub
'    End If
'
'auth_ErrorHandling:
'
'    auth_ErrorDescription = "An error occurred while logging in." & vbNewLine
'    If Err.Number <> 0 Then
'        If Err.Number - vbObjectError <> 11040 Then
'            auth_ErrorDescription = auth_ErrorDescription & _
'                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
'        End If
'    Else
'        auth_ErrorDescription = auth_ErrorDescription & "Login did not complete"
'    End If
'    auth_ErrorDescription = auth_ErrorDescription & Err.Description
'
'    WebHelpers.LogError auth_ErrorDescription, "GraphAuthenticator.Login", 11040 + vbObjectError
'    Err.Raise 11040 + vbObjectError, "GraphAuthenticator.Login", auth_ErrorDescription
'End Sub
'
Public Sub Login() ' If you have Access with the new Web browser control
    On Error GoTo auth_ErrorHandling
    Dim dtEnd As Date

    ' No need to login if API key, authorization code, or token have been set
    If Me.ApiKey <> "" Or Me.AuthorizationCode <> "" Or Me.Token <> "" Then
        Exit Sub
    End If

    Dim auth_Completed As Boolean

    ' Windows login uses new Edge Browser Control to automate retrieving authorization code for user
    Dim auth_ErrorDescription As String
    On Error GoTo auth_Cleanup

    auth_Completed = False

    Dim auth_IE As Object
    DoCmd.OpenForm "EdgeForm"
    Set auth_IE = Forms!EdgeForm.EdgeBrowser
    auth_IE.Navigate GetLoginUrl()

    ' Wait for login to complete
    dtEnd = DateAdd("s", WaitForLogin, Now)
    Do While Not auth_LoginIsComplete(auth_IE)
        DoEvents
        If Now > dtEnd Then
            Err.Raise 11042 + vbObjectError, "OAuthDialog", "Login error: Wait time exceeded."
            Exit Do
        End If
    Loop
    auth_Completed = True

    If auth_LoginIsDenied(auth_IE) Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login failed or was denied"
    ElseIf auth_LoginIsError(auth_IE) Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login error: " & auth_LoginExtractError(auth_IE)
    End If

    ' Success!
    Me.AuthorizationCode = auth_LoginExtractCode(auth_IE)

auth_Cleanup:

    If Not auth_IE Is Nothing Then: DoCmd.Close acForm, "EdgeForm"
    Set auth_IE = Nothing

    If Err.Number = 0 And auth_Completed Then
        WebHelpers.LogDebug "Login succeeded: " & Me.AuthorizationCode, "GraphAuthenticator.Login"
        Exit Sub
    End If

auth_ErrorHandling:

    auth_ErrorDescription = "An error occurred while logging in." & vbNewLine
    If Err.Number <> 0 Then
        If Err.Number - vbObjectError <> 11040 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
    Else
        auth_ErrorDescription = auth_ErrorDescription & "Login did not complete"
    End If
    auth_ErrorDescription = auth_ErrorDescription & Err.Description

    WebHelpers.LogError auth_ErrorDescription, "GraphAuthenticator.Login", 11040 + vbObjectError
    Err.Raise 11040 + vbObjectError, "GraphAuthenticator.Login", auth_ErrorDescription
End Sub

''
' Logout
''
Public Sub Logout()
    Me.AuthorizationCode = ""
    Me.Token = ""
End Sub

''
' Add scope to authorized scopes
'
' - To get a list of available scopes, visit https://learn.microsoft.com/en-us/graph/permissions-reference
'
' @example
' Auth.AddScope "mail.readwrite"
'
' @param {String} Scope
''
Public Sub AddScope(Scope As String)
    Dim auth_Scopes As Variant
    
    ' Prepare scopes array
    auth_Scopes = Me.Scopes
    If VBA.IsEmpty(auth_Scopes) Then
        ReDim auth_Scopes(0 To 0)
    Else
        ReDim Preserve auth_Scopes(0 To UBound(auth_Scopes) + 1)
    End If
    
    auth_Scopes(UBound(auth_Scopes)) = Scope
    Me.Scopes = auth_Scopes
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    If Me.Token = "" Then
        If Me.AuthorizationCode = "" Then
            Me.Login
        End If
        If Me.AuthorizationCode <> "" Then
            Me.Token = Me.GetToken(Client)
        Else
            Exit Sub
        End If
    End If

    Request.SetHeader "Authorization", "Bearer " & Me.Token
    Request.SetHeader "Content-Type", WebFormat.JSON
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    ' e.g. Handle 401 Unauthorized or other issues
End Sub

''
' Hook for updating http before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {WinHttpRequest} Http
''
Private Sub IWebAuthenticator_PrepareHttp(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Http As Object)
    ' e.g. Update option, headers, etc.
End Sub

''
' Hook for updating cURL before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {String} Curl
''
Private Sub IWebAuthenticator_PrepareCurl(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
    ' e.g. Add flags to cURL
End Sub

''
' Get token (for current AuthorizationCode)
'
' @internal
' @param {WebClient} Client
' @return {String}
''
Public Function GetToken(Client As WebClient) As String
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clone client (to avoid accidental interactions)
    Set auth_TokenClient = Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://login.microsoftonline.com/"
    
    ' Prepare token request
    auth_Request.Resource = Me.TenantID & "/oauth2/v2.0/token"
    auth_Request.Method = WebMethod.HttpPOST
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.JSON
    
    auth_Body.Add "client_id", Me.ClientId
    auth_Body.Add "scope", Me.Scopes(1)
    auth_Body.Add "code", Me.AuthorizationCode
    auth_Body.Add "redirect_uri", auth_RedirectUrl
    auth_Body.Add "grant_type", "authorization_code"
'    auth_Body.Add "client_secret", Me.ClientSecret
    Set auth_Request.body = auth_Body
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    If auth_Response.StatusCode = WebStatusCode.OK Then
        GetToken = auth_Response.Data("access_token")
    Else
        Err.Raise 11041 + vbObjectError, "GraphAuthenticator.GetToken", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    
    Set auth_TokenClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving token." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "GraphAuthenticator.GetToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "GraphAuthenticator.GetToken", auth_ErrorDescription
    End If
End Function

''
' Get login url for current scopes
'
' @internal
' @return {String}
''
Public Function GetLoginUrl() As String
    ' Use Request for Url helpers
    Dim auth_Request As New WebRequest
    auth_Request.Resource = AuthorizationUrl
    
    auth_Request.AddQuerystringParam "redirect_uri", auth_RedirectUrl
    auth_Request.AddQuerystringParam "client_id", Me.ClientId
    auth_Request.AddQuerystringParam "response_type", "code"
    auth_Request.AddQuerystringParam "access_type", "offline"
    auth_Request.AddQuerystringParam "approval_prompt", "force"
    
    If Not VBA.IsEmpty(Me.Scopes) Then
        auth_Request.AddQuerystringParam "scope", VBA.Join(Me.Scopes, " ")
    Else
        auth_Request.AddQuerystringParam "scope", ""
    End If
    
    GetLoginUrl = auth_Request.FormattedResource
    Set auth_Request = Nothing
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Function IEauth_LoginIsComplete(auth_IE As Object) As Boolean
    If Not auth_IE.Busy And auth_IE.ReadyState = 4 Then
        IEauth_LoginIsComplete = auth_LoginIsApproval(auth_IE) Or IEauth_LoginIsError(auth_IE)
    End If
End Function

Private Function IEauth_LoginIsDenied(auth_IE As Object) As Boolean
    Dim auth_Document As Object
    Dim auth_Element As Object

    If auth_LoginIsApproval(auth_IE) Then
        For Each auth_Element In auth_IE.Document.body.All
            If VBA.UCase(auth_Element.NodeName) = "P" And auth_Element.Id = "access_denied" Then
                IEauth_LoginIsDenied = True
                Exit Function
            End If
        Next auth_Element
    End If
End Function

Private Function IEauth_LoginIsError(auth_IE As Object) As Boolean
    DoEvents
    IEauth_LoginIsError = InStr(1, auth_IE.Document.body.innerHTML, "error") > 0
End Function

Private Function IEauth_LoginExtractError(auth_IE As Object) As String
    Dim auth_Element As Object

    For Each auth_Element In auth_IE.Document.body.All
        If auth_Element.Id = "exceptionMessageContainer" Then
            IEauth_LoginExtractError = auth_Element.innerHTML
            Exit Function
        End If
    Next auth_Element
End Function

Private Function auth_LoginIsApproval(auth_IE As Object) As Boolean
    auth_LoginIsApproval = InStr(auth_IE.LocationURL, "?code=") > 1
End Function

Private Function auth_LoginExtractCode(auth_IE As Object) As String
    Dim sRedirectURL As String
    
    sRedirectURL = auth_IE.LocationURL
    sRedirectURL = Mid(sRedirectURL, InStr(sRedirectURL, "code=") + 5)
    sRedirectURL = Left(sRedirectURL, InStr(sRedirectURL, "session_state=") - 2)
    auth_LoginExtractCode = sRedirectURL
End Function

Private Function auth_LoginIsComplete(auth_IE As Object) As Boolean
    If auth_IE.ReadyState = 4 Then
        auth_LoginIsComplete = auth_LoginIsApproval(auth_IE) Or auth_LoginIsError(auth_IE)
    End If
End Function

Private Function auth_LoginIsError(auth_IE As Object) As Boolean
    DoEvents
    If InStr(auth_IE.LocationURL, "scope") > 1 Then
        auth_LoginIsError = False
    Else
        auth_LoginIsError = InStr(1, auth_IE.RetrieveJavascriptValue("document.documentElement.innerHTML"), "error") > 0
    End If
End Function

Private Function auth_LoginIsDenied(auth_IE As Object) As Boolean
    Dim auth_Document As Object
    Dim auth_Element As Object

    If auth_LoginIsApproval(auth_IE) Then
'        For Each auth_Element In auth_IE.Object.Document.body.All
'            If VBA.UCase(auth_Element.NodeName) = "P" And auth_Element.Id = "access_denied" Then
'                auth_LoginIsDenied = True
'                Exit Function
'            End If
'        Next auth_Element
    End If
End Function

Private Function auth_LoginExtractError(auth_IE As Object) As String
    auth_LoginExtractError = auth_IE.RetrieveJavascriptValue("document.getElementById('exceptionMessageContainer').innerHTML")
End Function


Private Sub Class_Initialize()
    Me.Scopes = Array("")
End Sub
