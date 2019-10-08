VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WebClient v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' `WebClient` executes requests and handles response and is responsible for functionality shared between requests,
'  such as authentication, proxy configuration, and security.
'
' Usage:
'
' ```VB.net
' Dim Client As New WebClient
' Client.BaseUrl = "https://www.example.com/api/"
'
' Dim Auth As New HttpBasicAuthenticator
' Auth.Setup Username, Password
' Set Client.Authenticator = Auth
'
' Dim Request As New WebRequest
' Dim Response As WebResponse
' ' Setup WebRequest...
'
' Set Response = Client.Execute(Request)
' ' -> Uses Http Basic authentication and appends Request.Resource to BaseUrl
' ```
'
' Errors:
' 11010 / 80042b02 / -2147210494 - cURL error in Execute
' 11011 / 80042b03 / -2147210493 - Error in Execute
' 11012 / 80042b04 / -2147210492 - Error preparing http request
' 11013 / 80042b05 / -2147210491 - Error preparing cURL request
'
' @class WebClient
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit


Private Const web_DefaultTimeoutMs As Long = 5000

Private Const web_HttpRequest_SetCredentials_ForServer = 0
Private Const web_HttpRequest_SetCredentials_ForProxy = 1

Private Const web_HttpRequest_ProxySetting_Default = 0
Private Const web_HttpRequest_ProxySetting_PreConfig = 0
Private Const web_HttpRequest_ProxySetting_Direct = 1
Private Const web_HttpRequest_ProxySetting_Proxy = 2

Private Enum web_WinHttpRequestOption
    web_WinHttpRequestOption_UserAgentString = 0
    web_WinHttpRequestOption_URL = 1
    web_WinHttpRequestOption_URLCodePage = 2
    web_WinHttpRequestOption_EscapePercentInURL = 3
    web_WinHttpRequestOption_SslErrorIgnoreFlags = 4
    web_WinHttpRequestOption_SelectCertificate = 5
    web_WinHttpRequestOption_EnableRedirects = 6
    web_WinHttpRequestOption_UrlEscapeDisable = 7
    web_WinHttpRequestOption_UrlEscapeDisableQuery = 8
    web_WinHttpRequestOption_SecureProtocols = 9
    web_WinHttpRequestOption_EnableTracing = 10
    web_WinHttpRequestOption_RevertImpersonationOverSsl = 11
    web_WinHttpRequestOption_EnableHttpsToHttpRedirects = 12
    web_WinHttpRequestOption_EnablePassportAuthentication = 13
    web_WinHttpRequestOption_MaxAutomaticRedirects = 14
    web_WinHttpRequestOption_MaxResponseHeaderSize = 15
    web_WinHttpRequestOption_MaxResponseDrainSize = 16
    web_WinHttpRequestOption_EnableHttp1_1 = 17
    web_WinHttpRequestOption_EnableCertificateRevocationCheck = 18
End Enum

Private web_pProxyServer As String
Private web_pAutoProxyDomain As String


Public BaseUrl As String

Public Authenticator As IWebAuthenticator


Public TimeoutMs As Long

Public ConnectionTimeoutMs As Long

Public ProxyBypassList As String

Public ProxyUsername As String

Public ProxyPassword As String

Public EnableAutoProxy As Boolean

Public Insecure As Boolean

Public FollowRedirects As Boolean

Public Property Get ProxyServer() As String
    ProxyServer = web_pProxyServer
End Property
Public Property Let ProxyServer(Value As String)
    Me.EnableAutoProxy = False
    web_pProxyServer = Value
End Property


Public Function Execute(Request As WebRequest) As WebResponse
    Dim web_Http As Object
    Dim web_Response As New WebResponse

    On Error GoTo web_ErrorHandling

#If Mac Then
    Dim web_Curl As String
    Dim web_Result As ShellResult

    web_Curl = Me.PrepareCurlRequest(Request)
    web_Result = WebHelpers.ExecuteInShell(web_Curl)

    If web_Result.ExitCode > 0 Then
        Dim web_ErrorNumber As Long
        Dim web_ErrorMessage As String
        Dim web_ErrorDetails As String
        web_ErrorNumber = web_Result.ExitCode / 256

        Select Case web_ErrorNumber
        Case 1
            Err.Raise 12006 + &H30000 + vbObjectError, "The URL does not use a recognized protocol (1: CURLE_UNSUPPORTED_PROTOCOL)" & vbNewLine & _
                "URL: " & Me.GetFullUrl(Request) & vbNewLine & _
                "Protocol: " & WebHelpers.GetUrlParts(Me.GetFullUrl(Request))("Protocol")
        Case 3
            Err.Raise 12005 + &H30000 + vbObjectError, "The URL is invalid (3: CURLE_URL_MALFORMAT)" & _
                "URL: " & Me.GetFullUrl(Request)
        Case 5, 6
            If web_ErrorNumber = 5 Then
                web_ErrorDetails = "(5: CURLE_COULDNT_RESOLVE_PROXY)"
            Else
                web_ErrorDetails = "(6: CURLE_COULDNT_RESOLVE_HOST)"
            End If

            Err.Raise 12007 + &H30000 + vbObjectError, "WebClient.Execute", "The server name or address could not be resolved " & web_ErrorDetails
        Case 7
            Err.Raise 12029 + &H30000 + vbObjectError, "WebClient.Execute", "A connection with the server could not be established (7: CURLE_COULDNT_CONNECT)"
        Case 12, 28
            If web_ErrorNumber = 12 Then
                web_ErrorDetails = "(12: CURLE_FTP_ACCEPT_TIMEOUT)"
            Else
                web_ErrorDetails = "(28: CURLE_OPERATION_TIMEDOUT)"
            End If

            Err.Raise 12002 + &H30000 + vbObjectError, "WebClient.Execute", "The operation timed out " & web_ErrorDetails
        Case 47
            Err.Raise 12156 + &H30000 + vbObjectError, "WebClient.Execute", "Too many redirects (47: CURLE_TOO_MANY_REDIRECTS)"
        Case Else
            Err.Raise 11010 + vbObjectError, "WebClient.Execute", "An unknown cURL error occured, #" & web_ErrorNumber & vbNewLine & _
                "Find details at http://curl.haxx.se/libcurl/c/libcurl-errors.html"
        End Select
    End If

    web_Response.CreateFromCurl Me, Request, web_Result.Output

#Else
    Set web_Http = Me.PrepareHttpRequest(Request, False)

    web_Http.Send Request.Body
    'Do While Not web_Http.WaitForResponse(0.025)
    '    VBA.DoEvents
    'Loop

    web_Response.CreateFromHttp Me, Request, web_Http

#End If
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.AfterExecute Me, Request, web_Response
    End If

    Set web_Http = Nothing
    Set Execute = web_Response
    Exit Function

web_ErrorHandling:

    Set web_Http = Nothing
    Dim web_ErrorDescription As String

    Select Case Err.Number And 65535
    Case 12002, 12007, 12029
        web_Response.StatusCode = WebStatusCode.RequestTimeout
        web_Response.StatusDescription = "Request Timeout: " & Err.Description

        Set Execute = web_Response
        Err.Clear
    Case Else
        web_ErrorDescription = "An error occurred during execute" & vbNewLine & _
            Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

        Err.Raise 11011 + vbObjectError, "WebClient.Execute", web_ErrorDescription
    End Select
End Function

Public Function GetJson(Url As String, Optional Options As Dictionary) As WebResponse
    Dim web_Request As New WebRequest
    web_Request.CreateFromOptions Options
    web_Request.Resource = Url
    web_Request.Format = WebFormat.Json
    web_Request.Method = WebMethod.HttpGet

    Set GetJson = Me.Execute(web_Request)
End Function

Public Function PostJson(Url As String, Body As Variant, Optional Options As Dictionary) As WebResponse
    Dim web_Request As New WebRequest
    web_Request.CreateFromOptions Options
    web_Request.Resource = Url
    web_Request.Format = WebFormat.Json
    web_Request.Method = WebMethod.HttpPost
    If VBA.IsObject(Body) Then
        Set web_Request.Body = Body
    Else
        web_Request.Body = Body
    End If

    Set PostJson = Me.Execute(web_Request)
End Function

Public Sub SetProxy(ProxyServer As String, _
    Optional Username As String = "", Optional Password As String = "", Optional BypassList As String = "")

    Me.ProxyServer = ProxyServer
    Me.ProxyUsername = Username
    Me.ProxyPassword = Password
    Me.ProxyBypassList = BypassList
End Sub

Public Function GetFullUrl(Request As WebRequest) As String
    GetFullUrl = WebHelpers.JoinUrl(Me.BaseUrl, Request.FormattedResource)
End Function

Public Function PrepareHttpRequest(Request As WebRequest, Optional Async As Boolean = True) As Object
    Dim web_Http As Object
    Dim web_KeyValue As Dictionary

    On Error GoTo web_ErrorHandling

    Set web_Http = CreateObject("WinHttp.WinHttpRequest.5.1")

    web_BeforeExecute Request

    web_Http.Open WebHelpers.MethodToName(Request.Method), Me.GetFullUrl(Request), Async

    web_Http.SetTimeouts Me.ConnectionTimeoutMs, Me.ConnectionTimeoutMs, Me.ConnectionTimeoutMs, Me.TimeoutMs

    If Me.EnableAutoProxy Then
        web_LoadAutoProxy Request
    End If

    If Me.ProxyServer <> "" Then
        web_Http.SetProxy web_HttpRequest_ProxySetting_Proxy, Me.ProxyServer, Me.ProxyBypassList

        If Me.ProxyUsername <> "" Then
            web_Http.SetCredentials Me.ProxyUsername, Me.ProxyPassword, web_HttpRequest_SetCredentials_ForProxy
        End If
    Else
        web_Http.SetProxy web_HttpRequest_ProxySetting_PreConfig
    End If

    If Me.Insecure Then
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableCertificateRevocationCheck) = False
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_SslErrorIgnoreFlags) = 13056
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableHttpsToHttpRedirects) = True
    Else
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableCertificateRevocationCheck) = True
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_SslErrorIgnoreFlags) = 0
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableHttpsToHttpRedirects) = False
    End If

    web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableRedirects) = Me.FollowRedirects

    For Each web_KeyValue In Request.Headers
        web_Http.SetRequestHeader web_KeyValue.Item("Key"), web_KeyValue.Item("Value")
    Next web_KeyValue

    For Each web_KeyValue In Request.Cookies
        web_Http.SetRequestHeader "Cookie", web_KeyValue.Item("Key") & "=" & web_KeyValue.Item("Value")
    Next web_KeyValue

    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.PrepareHttp Me, Request, web_Http
    End If

    Set PrepareHttpRequest = web_Http
    Exit Function

web_ErrorHandling:

    Set web_Http = Nothing
    Err.Raise 11012 + vbObjectError, "WebClient.PrepareHttpRequest", _
        "An error occurred while preparing http request" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
End Function

Public Function PrepareCurlRequest(Request As WebRequest) As String
    Dim web_Curl As String
    Dim web_KeyValue As Dictionary
    Dim web_CookieString As String

    On Error GoTo web_ErrorHandling

    web_Curl = "curl -i"

    web_BeforeExecute Request

    web_Curl = web_Curl & " --connect-timeout " & Me.ConnectionTimeoutMs / 1000
    web_Curl = web_Curl & " --max-time " & 3 * Me.TimeoutMs / 1000

    If Me.ProxyServer <> "" Then
        web_Curl = web_Curl & " --proxy " & Me.ProxyServer

        If Me.ProxyBypassList <> "" Then
            web_Curl = web_Curl & " --noproxy " & Me.ProxyBypassList
        End If
        If Me.ProxyUsername <> "" Then
            web_Curl = web_Curl & " --proxy-user " & Me.ProxyUsername & ":" & Me.ProxyPassword
        End If
    End If

    If Me.Insecure Then
        web_Curl = web_Curl & " --insecure"
    End If

    If Me.FollowRedirects Then
        web_Curl = web_Curl & " --location"
    End If
    
    If Request.AcceptEncoding <> "identity" Then
        web_Curl = web_Curl & " --compressed"
    End If

    For Each web_KeyValue In Request.Headers
        web_Curl = web_Curl & " -H '" & web_KeyValue.Item("Key") & ": " & web_KeyValue.Item("Value") & "'"
    Next web_KeyValue

    For Each web_KeyValue In Request.Cookies
        web_CookieString = web_CookieString & web_KeyValue.Item("Key") & "=" & web_KeyValue.Item("Value") & ";"
    Next web_KeyValue
    If web_CookieString <> "" Then
        web_Curl = web_Curl & " --cookie '" & web_CookieString & "'"
    End If

    web_Curl = web_Curl & " -X " & WebHelpers.MethodToName(Request.Method)
    web_Curl = web_Curl & " -d '" & Request.Body & "'"
    web_Curl = web_Curl & " '" & Me.GetFullUrl(Request) & "'"

    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.PrepareCurl Me, Request, web_Curl
    End If

    PrepareCurlRequest = web_Curl
    Exit Function

web_ErrorHandling:

    Err.Raise 11013 + vbObjectError, "WebClient.PrepareCURLRequest", _
        "An error occurred while preparing cURL request" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
End Function

Public Function Clone() As WebClient
    Set Clone = New WebClient
    Clone.BaseUrl = Me.BaseUrl
    Clone.ProxyServer = Me.ProxyServer
    Clone.ProxyBypassList = Me.ProxyBypassList
    Clone.ProxyUsername = Me.ProxyUsername
    Clone.ProxyPassword = Me.ProxyPassword
    Clone.EnableAutoProxy = Me.EnableAutoProxy
    Clone.TimeoutMs = Me.TimeoutMs
    Clone.ConnectionTimeoutMs = Me.ConnectionTimeoutMs
    Clone.Insecure = Me.Insecure
    Set Clone.Authenticator = Me.Authenticator
End Function


Private Sub web_BeforeExecute(web_Request As WebRequest)
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, web_Request
    End If

    web_Request.Prepare
End Sub

Private Sub web_LoadAutoProxy(web_Request As WebRequest)
#If Win32 Or Win64 Then
    On Error GoTo web_ErrorHandling

    Dim web_Parts As Dictionary
    Dim web_Domain As String
    Dim web_ProxyServer As String
    Dim web_ProxyBypassList As String

    Set web_Parts = WebHelpers.GetUrlParts(Me.GetFullUrl(web_Request))
    web_Domain = VBA.IIf(web_Parts.Item("Protocol") <> "", web_Parts.Item("Protocol") & "://", "") & _
        web_Parts.Item("Host") & ":" & web_Parts.Item("Port")

    If web_Domain <> web_pAutoProxyDomain Then
        web_pAutoProxyDomain = web_Domain

        WebHelpers.GetAutoProxy web_Domain, web_ProxyServer, web_ProxyBypassList

        WebHelpers.LogDebug "Loaded auto-proxy for " & web_Domain & ":" & vbNewLine & _
            "Server = " & web_ProxyServer & vbNewLine & _
            "Bypass List = " & web_ProxyBypassList

        web_pProxyServer = web_ProxyServer
        Me.ProxyBypassList = web_ProxyBypassList
    End If

    Exit Sub

web_ErrorHandling:

    WebHelpers.LogError "An error occurred while loading auto-proxy" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description, _
        "WebClient.LoadAutoProxy", Err.Number
#End If
End Sub

Private Sub Class_Initialize()
    Me.TimeoutMs = web_DefaultTimeoutMs
    Me.ConnectionTimeoutMs = web_DefaultTimeoutMs
    Me.EnableAutoProxy = False
    Me.Insecure = False
    Me.FollowRedirects = True
End Sub
