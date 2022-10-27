VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WebRequest v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' `WebRequest` is used to create detailed requests
' (including formatting, querystrings, headers, cookies, and much more).
'
' Usage:
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "users/{Id}"
'
' Request.Method = WebMethod.HttpPut
' Request.RequestFormat = WebFormat.UrlEncoded
' Request.ResponseFormat = WebFormat.Json
'
' Dim Body As New Dictionary
' Body.Add "name", "Tim"
' Body.Add "project", "VBA-Web"
' Set Request.Body = Body
'
' Request.AddUrlSegment "Id", 123
' Request.AddQuerystringParam "api_key", "abcd"
' Request.AddHeader "Authorization", "Token ..."
'
' ' -> PUT (Client.BaseUrl)users/123?api_key=abcd
' '    Authorization: Token ...
' '
' '    name=Tim&project=VBA-Web
' ```
'
' Errors:
' 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
'
' @class WebRequest
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit


Private web_pRequestFormat As WebFormat
Private web_pResponseFormat As WebFormat
Private web_pCustomRequestFormat As String
Private web_pCustomResponseFormat As String
Private web_pBody As Variant
Private web_pConvertedBody As Variant
Private web_pContentType As String
Private web_pAccept As String
Private web_pAcceptEncoding As String
Private web_pContentLength As Long
Private web_pId As String


Public Resource As String

Public Method As WebMethod

Public Headers As Collection

Public QuerystringParams As Collection

Public UrlSegments As Dictionary

Public Cookies As Collection

Public UserAgent As String

Public Property Get Format() As WebFormat
    Format = RequestFormat
End Property
Public Property Let Format(Value As WebFormat)
    Me.RequestFormat = Value
    Me.ResponseFormat = Value
End Property

Public Property Get RequestFormat() As WebFormat
    RequestFormat = web_pRequestFormat
End Property
Public Property Let RequestFormat(Value As WebFormat)
    If Value <> web_pRequestFormat Then
        web_pRequestFormat = Value

        web_pConvertedBody = Empty
    End If
End Property

Public Property Get ResponseFormat() As WebFormat
    ResponseFormat = web_pResponseFormat
End Property
Public Property Let ResponseFormat(Value As WebFormat)
    If Value <> web_pResponseFormat Then
        web_pResponseFormat = Value

        web_pConvertedBody = Empty
    End If
End Property

Public Property Get CustomRequestFormat() As String
    CustomRequestFormat = web_pCustomRequestFormat
End Property
Public Property Let CustomRequestFormat(Value As String)
    If Value <> web_pCustomRequestFormat Then
        web_pCustomRequestFormat = Value

        web_pConvertedBody = Empty

        If Value <> "" Then
            web_pRequestFormat = WebFormat.Custom
        End If
    End If
End Property

Public Property Get CustomResponseFormat() As String
    CustomResponseFormat = web_pCustomResponseFormat
End Property
Public Property Let CustomResponseFormat(Value As String)
    If Value <> web_pCustomResponseFormat Then
        web_pCustomResponseFormat = Value

        web_pConvertedBody = Empty

        If Value <> "" Then
            ResponseFormat = WebFormat.Custom
        End If
    End If
End Property

Public Property Get ContentType() As String
    If web_pContentType <> "" Then
        ContentType = web_pContentType
    Else
        ContentType = WebHelpers.FormatToMediaType(Me.RequestFormat, Me.CustomRequestFormat)
    End If
End Property
Public Property Let ContentType(Value As String)
    web_pContentType = Value
End Property

Public Property Get Accept() As String
    If web_pAccept <> "" Then
        Accept = web_pAccept
    Else
        Accept = WebHelpers.FormatToMediaType(Me.ResponseFormat, Me.CustomResponseFormat)
    End If
End Property
Public Property Let Accept(Value As String)
    web_pAccept = Value
End Property

Public Property Get AcceptEncoding() As String
    If web_pAcceptEncoding <> "" Then
        AcceptEncoding = web_pAcceptEncoding
    Else
        AcceptEncoding = "identity"
    End If
End Property
Public Property Let AcceptEncoding(Value As String)
    web_pAcceptEncoding = Value
End Property

Public Property Get ContentLength() As Long
    If web_pContentLength >= 0 Then
        ContentLength = web_pContentLength
    Else
        ContentLength = Len(Me.Body)
    End If
End Property
Public Property Let ContentLength(Value As Long)
    web_pContentLength = Value
End Property

Public Property Get Body() As Variant
    If Not VBA.IsEmpty(web_pBody) Then
        If VBA.VarType(web_pBody) = vbString Then
            Body = web_pBody
        ElseIf IsEmpty(web_pConvertedBody) Then
            Body = WebHelpers.ConvertToFormat(web_pBody, Me.RequestFormat, Me.CustomRequestFormat)
            web_pConvertedBody = Body
        Else
            Body = web_pConvertedBody
        End If
    End If
End Property
Public Property Let Body(Value As Variant)
    web_pConvertedBody = Empty
    web_pBody = Value
End Property
Public Property Set Body(Value As Variant)
    web_pConvertedBody = Empty
    Set web_pBody = Value
End Property

Public Property Get FormattedResource() As String
    Dim web_Segment As Variant
    Dim web_Encoding As UrlEncodingMode

    FormattedResource = Me.Resource

    For Each web_Segment In Me.UrlSegments.Keys
        FormattedResource = VBA.Replace(FormattedResource, "{" & web_Segment & "}", WebHelpers.UrlEncode(Me.UrlSegments.Item(web_Segment)))
    Next web_Segment

    If Me.QuerystringParams.Count > 0 Then
        If VBA.InStr(FormattedResource, "?") <= 0 Then
            FormattedResource = FormattedResource & "?"
        Else
            FormattedResource = FormattedResource & "&"
        End If

        If Me.RequestFormat = WebFormat.FormUrlEncoded Then
            web_Encoding = UrlEncodingMode.FormUrlEncoding
        Else
            web_Encoding = UrlEncodingMode.QueryUrlEncoding
        End If
        FormattedResource = FormattedResource & WebHelpers.ConvertToUrlEncoded(Me.QuerystringParams, EncodingMode:=web_Encoding)
    End If
End Property

Public Property Get Id() As String
    If web_pId = "" Then: web_pId = WebHelpers.CreateNonce
    Id = web_pId
End Property


Public Sub AddHeader(Key As String, Value As Variant)
    Me.Headers.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

Public Sub SetHeader(Key As String, Value As Variant)
    WebHelpers.AddOrReplaceInKeyValues Me.Headers, Key, Value
End Sub

Public Sub AddUrlSegment(Segment As String, Value As Variant)
    Me.UrlSegments.Item(Segment) = Value
End Sub

Public Sub AddQuerystringParam(Key As String, Value As Variant)
    Me.QuerystringParams.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

Public Sub AddCookie(Key As String, Value As Variant)
    Me.Cookies.Add WebHelpers.CreateKeyValue( _
        web_EncodeCookieName(Key), _
        WebHelpers.UrlEncode(Value, EncodingMode:=UrlEncodingMode.CookieUrlEncoding) _
    )
End Sub

Public Sub AddBodyParameter(Key As Variant, Value As Variant)
    If VBA.IsEmpty(web_pBody) Then
        Set web_pBody = New Dictionary
    ElseIf Not TypeOf web_pBody Is Dictionary Then
        Dim web_ErrorDescription As String
        web_ErrorDescription = "Cannot add body parameter to non-Dictionary Body (existing Body must be of type Dictionary)"

        WebHelpers.LogError web_ErrorDescription, "WebRequest.AddBodyParameter", 11020 + vbObjectError
        Err.Raise 11020 + vbObjectError, "WebRequest.AddBodyParameter", web_ErrorDescription
    End If

    If VBA.IsObject(Value) Then
        Set web_pBody.Item(Key) = Value
    Else
        web_pBody.Item(Key) = Value
    End If

    web_pConvertedBody = Empty
End Sub

Public Sub Prepare()
    SetHeader "User-Agent", Me.UserAgent
    SetHeader "Accept", Me.Accept
    SetHeader "Accept-Encoding", Me.AcceptEncoding
    If Me.Method <> WebMethod.HttpGet Or Me.ContentLength > 0 Then
        SetHeader "Content-Type", Me.ContentType
        SetHeader "Content-Length", VBA.CStr(Me.ContentLength)
    End If
End Sub

Public Function Clone() As WebRequest
    Set Clone = New WebRequest

    Clone.Resource = Me.Resource
    Clone.Method = Me.Method
    Clone.UserAgent = Me.UserAgent
    Clone.Accept = web_pAccept
    Clone.ContentType = web_pContentType
    Clone.ContentLength = web_pContentLength
    Clone.RequestFormat = Me.RequestFormat
    Clone.ResponseFormat = Me.ResponseFormat
    Clone.CustomRequestFormat = Me.CustomRequestFormat
    Clone.CustomResponseFormat = Me.CustomResponseFormat

    Set Clone.Headers = WebHelpers.CloneCollection(Me.Headers)
    Set Clone.QuerystringParams = WebHelpers.CloneCollection(Me.QuerystringParams)
    Set Clone.UrlSegments = WebHelpers.CloneDictionary(Me.UrlSegments)
    Set Clone.Cookies = WebHelpers.CloneCollection(Me.Cookies)

    If VBA.IsObject(web_pBody) Then
        Set Clone.Body = web_pBody
    Else
        Clone.Body = web_pBody
    End If
End Function

Public Sub CreateFromOptions(Options As Dictionary)
    If Not Options Is Nothing Then
        If Options.Exists("Headers") Then
            Set Me.Headers = Options.Item("Headers")
        End If
        If Options.Exists("Cookies") Then
            Set Me.Cookies = Options.Item("Cookies")
        End If
        If Options.Exists("QuerystringParams") Then
            Set Me.QuerystringParams = Options.Item("QuerystringParams")
        End If
        If Options.Exists("UrlSegments") Then
            Set Me.UrlSegments = Options.Item("UrlSegments")
        End If
    End If
End Sub


Private Function web_EncodeCookieName(web_CookieName As Variant) As String
    Dim web_CookieVal As String
    Dim web_StringLen As Long

    web_CookieVal = VBA.CStr(web_CookieName)
    web_StringLen = VBA.Len(web_CookieVal)

    If web_StringLen > 0 Then
        Dim web_Result() As String
        Dim web_i As Long
        Dim web_CharCode As Integer
        Dim web_Char As String
        ReDim web_Result(web_StringLen)


        For web_i = 1 To web_StringLen
            web_Char = VBA.Mid$(web_CookieVal, web_i, 1)
            web_CharCode = VBA.Asc(web_Char)

            Select Case web_CharCode
                Case 65 To 90, 97 To 122
                    web_Result(web_i) = web_Char
                Case 48 To 57
                    web_Result(web_i) = web_Char
                Case 33, 35, 36, 38, 39, 42, 43, 45, 46, 94, 95, 96, 124, 126
                    web_Result(web_i) = web_Char

                Case 0 To 15
                    web_Result(web_i) = "%0" & VBA.Hex(web_CharCode)
                Case Else
                    web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
            End Select
        Next web_i

        web_EncodeCookieName = VBA.Join$(web_Result, "")
    End If
End Function

Private Sub Class_Initialize()
    Me.RequestFormat = WebFormat.Json
    Me.ResponseFormat = WebFormat.Json
    Me.UserAgent = WebUserAgent

    Set Me.Headers = New Collection
    Set Me.QuerystringParams = New Collection
    Set Me.UrlSegments = New Dictionary
    Set Me.Cookies = New Collection
    Me.ContentLength = -1
End Sub
