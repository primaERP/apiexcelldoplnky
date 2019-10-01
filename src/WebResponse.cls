VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WebResponse v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Wrapper for http/cURL responses that includes parsed Data based on WebRequest.ResponseFormat.
'
' Usage:
' ```VB.net
' Dim Response As WebResponse
' Set Response = Client.Execute(Request)
'
' If Response.StatusCode = WebStatusCode.Ok Then
'   ' Response.Headers, Response.Cookies
'   ' Response.Data -> Parsed Response.Content based on Request.ResponseFormat
'   ' Response.Body -> Raw response bytes
' Else
'   Debug.Print "Error: " & Response.StatusCode & " - " & Response.Content
' End If
' ```
'
' Errors:
' 11030 / 80042b16 / -2147210474 - Error creating from http
' 11031 / 80042b17 / -2147210473 - Error creating from cURL
' 11032 / 80042b18 / -2147210472 - Error extracting headers
'
' @class WebResponse
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private web_CrLf As String


Public StatusCode As WebStatusCode

Public StatusDescription As String

Public Content As String

Public Body As Variant

Public Data As Object

Public Headers As Collection

Public Cookies As Collection


Public Sub Update(Updated As WebResponse)
    Me.StatusCode = Updated.StatusCode
    Me.StatusDescription = Updated.StatusDescription
    Me.Content = Updated.Content
    Me.Body = Updated.Body
    Set Me.Headers = Updated.Headers
    Set Me.Cookies = Updated.Cookies
    Set Me.Data = Updated.Data
End Sub

Public Sub CreateFromHttp(Client As WebClient, Request As WebRequest, Http As Object)
    On Error GoTo web_ErrorHandling

    Me.StatusCode = Http.Status
    Me.StatusDescription = Http.StatusText
    Me.Content = Http.responseText
    Me.Body = Http.ResponseBody

    web_LoadValues Http.GetAllResponseHeaders, Me.Content, Me.Body, Request

    Exit Sub

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while creating response from http" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    Err.Raise 11030 + vbObjectError, "WebResponse.CreateFromHttp", web_ErrorDescription
End Sub

Public Sub CreateFromCurl(Client As WebClient, Request As WebRequest, Result As String)
    On Error GoTo web_ErrorHandling

    Dim web_Lines() As String

    web_Lines = VBA.Split(Result, web_CrLf)

    Me.StatusCode = web_ExtractStatusFromCurlResponse(web_Lines)
    Me.StatusDescription = web_ExtractStatusTextFromCurlResponse(web_Lines)
    Me.Content = web_ExtractResponseTextFromCurlResponse(web_Lines)
    Me.Body = WebHelpers.StringToAnsiBytes(Me.Content)

    web_LoadValues web_ExtractHeadersFromCurlResponse(web_Lines), Me.Content, Me.Body, Request

    Exit Sub

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while creating response from cURL" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    Err.Raise 11031 + vbObjectError, "WebResponse.CreateFromCurl", web_ErrorDescription
End Sub

Public Function ExtractHeaders(ResponseHeaders As String) As Collection
    On Error GoTo web_ErrorHandling

    Dim web_Lines As Variant
    Dim web_i As Integer
    Dim web_Headers As New Collection
    Dim web_Header As Dictionary
    Dim web_ColonPosition As Long
    Dim web_Multiline As Boolean

    web_Lines = VBA.Split(ResponseHeaders, web_CrLf)

    For web_i = LBound(web_Lines) To (UBound(web_Lines) + 1)
        If web_i > UBound(web_Lines) Then
            web_Headers.Add web_Header
        ElseIf web_Lines(web_i) <> "" Then
            web_ColonPosition = VBA.InStr(1, web_Lines(web_i), ":")
            If web_ColonPosition = 0 And Not web_Header Is Nothing Then
                web_Multiline = True
            ElseIf web_Multiline Then
                web_Multiline = False
                web_Headers.Add web_Header
            ElseIf Not web_Header Is Nothing Then
                web_Headers.Add web_Header
            End If

            If Not web_Multiline Then
                Set web_Header = WebHelpers.CreateKeyValue( _
                    Key:=VBA.Trim(VBA.Mid$(web_Lines(web_i), 1, web_ColonPosition - 1)), _
                    Value:=VBA.Trim(VBA.Mid$(web_Lines(web_i), web_ColonPosition + 1, VBA.Len(web_Lines(web_i)))) _
                )
            Else
                web_Header.Item("Value") = web_Header.Item("Value") & web_CrLf & web_Lines(web_i)
            End If
        End If
    Next web_i

    Set ExtractHeaders = web_Headers
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while extracting headers" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    Err.Raise 11032 + vbObjectError, "WebResponse.CreateFromCurl", web_ErrorDescription
End Function

Public Function ExtractCookies(Headers As Collection) As Collection
    Dim web_Header As Dictionary
    Dim web_Cookie As String
    Dim web_Key As String
    Dim web_Value As String
    Dim web_Cookies As New Collection

    For Each web_Header In Headers
        If web_Header.Item("Key") = "Set-Cookie" Then
            web_Cookie = web_Header.Item("Value")
            If VBA.InStr(1, web_Cookie, "=") > 0 Then
                web_Key = VBA.Mid$(web_Cookie, 1, VBA.InStr(1, web_Cookie, "=") - 1)
                web_Value = VBA.Mid$(web_Cookie, VBA.InStr(1, web_Cookie, "=") + 1, VBA.Len(web_Cookie))

                If VBA.InStr(1, web_Value, ";") > 0 Then
                    web_Value = VBA.Mid$(web_Value, 1, VBA.InStr(1, web_Value, ";") - 1)
                End If

                If VBA.Left$(web_Value, 1) = """" Then
                    web_Value = VBA.Mid$(web_Value, 2, VBA.Len(web_Value) - 2)
                End If

                web_Cookies.Add WebHelpers.CreateKeyValue(web_Key, WebHelpers.UrlDecode(web_Value, PlusAsSpace:=False, EncodingMode:=UrlEncodingMode.CookieUrlEncoding))
            Else

            End If
        End If
    Next web_Header

    Set ExtractCookies = web_Cookies
End Function


Private Sub web_LoadValues(web_Headers As String, web_Content As String, web_Body As Variant, web_Request As WebRequest)
    If web_Request.ResponseFormat <> WebFormat.PlainText Then
        On Error Resume Next
        Set Me.Data = _
            WebHelpers.ParseByFormat(web_Content, web_Request.ResponseFormat, web_Request.CustomResponseFormat, web_Body)

        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo 0
    End If

    Set Me.Headers = ExtractHeaders(web_Headers)

    Set Me.Cookies = ExtractCookies(Me.Headers)
End Sub

Private Function web_ExtractStatusFromCurlResponse(web_CurlResponseLines() As String) As Long
    Dim web_StatusLineParts() As String

    web_StatusLineParts = VBA.Split(web_CurlResponseLines(web_FindStatusLine(web_CurlResponseLines)), " ")
    web_ExtractStatusFromCurlResponse = VBA.CLng(web_StatusLineParts(1))
End Function

Private Function web_ExtractStatusTextFromCurlResponse(web_CurlResponseLines() As String) As String
    Dim web_StatusLineParts() As String
    Dim web_i As Long
    Dim web_StatusText As String

    web_StatusLineParts = VBA.Split(web_CurlResponseLines(web_FindStatusLine(web_CurlResponseLines)), " ", 3)
    web_ExtractStatusTextFromCurlResponse = web_StatusLineParts(2)
End Function

Private Function web_ExtractHeadersFromCurlResponse(web_CurlResponseLines() As String) As String
    Dim web_StatusLineIndex As Long
    Dim web_BlankLineIndex As Long
    Dim web_HeaderLines() As String
    Dim web_WriteIndex As Long
    Dim web_ReadIndex As Long

    web_StatusLineIndex = web_FindStatusLine(web_CurlResponseLines)
    web_BlankLineIndex = web_FindBlankLine(web_CurlResponseLines)

    ReDim web_HeaderLines(0 To web_BlankLineIndex - 2 - web_StatusLineIndex)

    web_WriteIndex = 0
    For web_ReadIndex = (web_StatusLineIndex + 1) To web_BlankLineIndex - 1
        web_HeaderLines(web_WriteIndex) = web_CurlResponseLines(web_ReadIndex)
        web_WriteIndex = web_WriteIndex + 1
    Next web_ReadIndex

    web_ExtractHeadersFromCurlResponse = VBA.Join$(web_HeaderLines, web_CrLf)
End Function

Private Function web_ExtractResponseTextFromCurlResponse(web_CurlResponseLines() As String) As String
    Dim web_BlankLineIndex As Long
    Dim web_BodyLines() As String
    Dim web_WriteIndex As Long
    Dim web_ReadIndex As Long

    web_BlankLineIndex = web_FindBlankLine(web_CurlResponseLines)

    ReDim web_BodyLines(0 To UBound(web_CurlResponseLines) - web_BlankLineIndex - 1)

    web_WriteIndex = 0
    For web_ReadIndex = web_BlankLineIndex + 1 To UBound(web_CurlResponseLines)
        web_BodyLines(web_WriteIndex) = web_CurlResponseLines(web_ReadIndex)
        web_WriteIndex = web_WriteIndex + 1
    Next web_ReadIndex

    web_ExtractResponseTextFromCurlResponse = VBA.Join$(web_BodyLines, web_CrLf)
End Function

Private Function web_FindStatusLine(web_CurlResponseLines() As String) As Long
    For web_FindStatusLine = LBound(web_CurlResponseLines) To UBound(web_CurlResponseLines)
        If VBA.Trim$(web_CurlResponseLines(web_FindStatusLine)) <> "" Then
            If VBA.Split(web_CurlResponseLines(web_FindStatusLine), " ")(1) <> "100" Then
                Exit Function
            End If
        End If
    Next web_FindStatusLine
End Function

Private Function web_FindBlankLine(web_CurlResponseLines() As String) As Long
    For web_FindBlankLine = (web_FindStatusLine(web_CurlResponseLines) + 1) To UBound(web_CurlResponseLines)
        If VBA.Trim$(web_CurlResponseLines(web_FindBlankLine)) = "" Then
            Exit Function
        End If
    Next web_FindBlankLine
End Function

Private Sub Class_Initialize()
    web_CrLf = VBA.Chr$(13) & VBA.Chr$(10)

    Set Headers = New Collection
    Set Cookies = New Collection
End Sub
