Attribute VB_Name = "WebHelpers"
''
' WebHelpers v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Contains general-purpose helpers that are used throughout VBA-Web. Includes:
'
' - Logging
' - Converters and encoding
' - Url handling
' - Object/Dictionary/Collection/Array helpers
' - Request preparation / handling
' - Timing
' - Mac
' - Cryptography
' - Converters (JSON, XML, Url-Encoded)
'
' Errors:
' 11000 - Error during parsing
' 11001 - Error during conversion
' 11002 - No matching converter has been registered
' 11003 - Error while getting url parts
' 11099 - XML format is not currently supported
'
' @module WebHelpers
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Public Const WebUserAgent As String = "VBA-Web v4.1.6 (https://github.com/VBA-tools/VBA-Web)"

Public Type ShellResult
    Output As String
    ExitCode As Long
End Type

Private Type json_Options
    UseDoubleForLargeNumbers As Boolean
    AllowUnquotedKeys As Boolean
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

Public Enum WebMethod
    HttpGet = 0
    HttpPost = 1
    HttpPut = 2
    HttpDelete = 3
    HttpPatch = 4
    HttpHead = 5
End Enum

Public Enum WebFormat
    PlainText = 0
    Json = 1
    FormUrlEncoded = 2
    Xml = 3
    Custom = 9
End Enum

Public Enum WebStatusCode
    Ok = 200
    Created = 201
    NoContent = 204
    NotModified = 304
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
    NotFound = 404
    RequestTimeout = 408
    UnsupportedMediaType = 415
    InternalServerError = 500
    BadGateway = 502
    ServiceUnavailable = 503
    GatewayTimeout = 504
End Enum

Public Enum UrlEncodingMode
    StrictUrlEncoding
    FormUrlEncoding
    QueryUrlEncoding
    CookieUrlEncoding
    PathUrlEncoding
End Enum

#If Mac Then
#If VBA7 Then
Private Declare PtrSafe Function web_popen Lib "/usr/lib/libc.dylib" Alias "popen" (ByVal web_Command As String, ByVal web_Mode As String) As LongPtr
Private Declare PtrSafe Function web_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" (ByVal web_File As LongPtr) As LongPtr
Private Declare PtrSafe Function web_fread Lib "/usr/lib/libc.dylib" Alias "fread" (ByVal web_OutStr As String, ByVal web_Size As LongPtr, ByVal web_Items As LongPtr, ByVal web_Stream As LongPtr) As LongPtr
Private Declare PtrSafe Function web_feof Lib "/usr/lib/libc.dylib" Alias "feof" (ByVal web_File As LongPtr) As LongPtr
#Else
Private Declare Function web_popen Lib "libc.dylib" Alias "popen" (ByVal web_Command As String, ByVal web_Mode As String) As Long
Private Declare Function web_pclose Lib "libc.dylib" Alias "pclose" (ByVal web_File As Long) As Long
Private Declare Function web_fread Lib "libc.dylib" Alias "fread" (ByVal web_OutStr As String, ByVal web_Size As Long, ByVal web_Items As Long, ByVal web_Stream As Long) As Long
Private Declare Function web_feof Lib "libc.dylib" Alias "feof" (ByVal web_File As Long) As Long
#End If
#End If

Public Function MethodToName(Method As WebMethod) As String
    Select Case Method
    Case WebMethod.HttpDelete
        MethodToName = "DELETE"
    Case WebMethod.HttpPut
        MethodToName = "PUT"
    Case WebMethod.HttpPatch
        MethodToName = "PATCH"
    Case WebMethod.HttpPost
        MethodToName = "POST"
    Case WebMethod.HttpGet
        MethodToName = "GET"
    Case WebMethod.HttpHead
        MethodToName = "HEAD"
    End Select
End Function

Public Function Obfuscate(Secure As String, Optional Character As String = "*") As String
    Obfuscate = VBA.String$(VBA.Len(Secure), Character)
End Function

Public Function Base64Encode(Text As String) As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl base64"
    Base64Encode = ExecuteInShell(web_Command).Output
#Else
    Dim web_Bytes() As Byte

    web_Bytes = VBA.StrConv(Text, vbFromUnicode)
    Base64Encode = web_AnsiBytesToBase64(web_Bytes)
#End If
    Base64Encode = VBA.Replace$(Base64Encode, vbLf, "")
End Function


#If Mac Then
#Else
Private Function web_AnsiBytesToBase64(web_Bytes() As Byte)
    Dim web_XmlObj As Object
    Dim web_Node As Object

    Set web_XmlObj = CreateObject("MSXML2.DOMDocument")
    Set web_Node = web_XmlObj.createElement("b64")

    web_Node.DataType = "bin.base64"
    web_Node.nodeTypedValue = web_Bytes
    web_AnsiBytesToBase64 = web_Node.Text

    Set web_Node = Nothing
    Set web_XmlObj = Nothing
End Function

Private Function web_AnsiBytesToHex(web_Bytes() As Byte)
    Dim web_i As Long
    For web_i = LBound(web_Bytes) To UBound(web_Bytes)
        web_AnsiBytesToHex = web_AnsiBytesToHex & VBA.LCase$(VBA.Right$("0" & VBA.Hex$(web_Bytes(web_i)), 2))
    Next web_i
End Function
#End If

Public Sub AddOrReplaceInKeyValues(KeyValues As Collection, Key As Variant, Value As Variant)
    Dim web_KeyValue As Dictionary
    Dim web_Index As Long
    Dim web_NewKeyValue As Dictionary

    Set web_NewKeyValue = CreateKeyValue(CStr(Key), Value)

    web_Index = 1
    For Each web_KeyValue In KeyValues
        If web_KeyValue.Item("Key") = Key Then
            KeyValues.Remove web_Index

            If KeyValues.Count = 0 Then
                KeyValues.Add web_NewKeyValue
            ElseIf web_Index > KeyValues.Count Then
                KeyValues.Add web_NewKeyValue, After:=web_Index - 1
            Else
                KeyValues.Add web_NewKeyValue, Before:=web_Index
            End If
            Exit Sub
        End If

        web_Index = web_Index + 1
    Next web_KeyValue

    KeyValues.Add web_NewKeyValue
End Sub

Public Function CreateKeyValue(Key As String, Value As Variant) As Dictionary
    Dim web_KeyValue As New Dictionary

    web_KeyValue.Item("Key") = Key
    web_KeyValue.Item("Value") = Value
    Set CreateKeyValue = web_KeyValue
End Function

Public Function FormatToMediaType(Format As WebFormat, Optional CustomFormat As String) As String
    Select Case Format
    Case WebFormat.FormUrlEncoded
        FormatToMediaType = "application/x-www-form-urlencoded;charset=UTF-8"
    Case WebFormat.Json
        FormatToMediaType = "application/json"
    Case WebFormat.Xml
        FormatToMediaType = "application/xml"
    Case WebFormat.Custom
        FormatToMediaType = web_GetConverter(CustomFormat)("MediaType")
    Case Else
        FormatToMediaType = "text/plain"
    End Select
End Function

Private Function web_GetConverter(web_CustomFormat As String) As Dictionary
    If web_pConverters.Exists(web_CustomFormat) Then
        Set web_GetConverter = web_pConverters(web_CustomFormat)
    Else
        LogError "No matching converter has been registered for custom format: " & web_CustomFormat, _
            "WebHelpers.web_GetConverter", 11002
        Err.Raise 11002, "WebHelpers.web_GetConverter", _
            "No matching converter has been registered for custom format: " & web_CustomFormat
    End If
End Function

Public Function ConvertToFormat(Obj As Variant, Format As WebFormat, Optional CustomFormat As String = "") As Variant
    On Error GoTo web_ErrorHandling

    Select Case Format
    Case WebFormat.Json
        ConvertToFormat = ConvertToJson(Obj)
    Case WebFormat.FormUrlEncoded
        ConvertToFormat = ConvertToUrlEncoded(Obj)
    Case WebFormat.Xml
        ConvertToFormat = ConvertToXml(Obj)
    Case WebFormat.Custom
#If EnableCustomFormatting Then
        Dim web_Converter As Dictionary
        Dim web_Callback As String

        Set web_Converter = web_GetConverter(CustomFormat)
        web_Callback = web_Converter.Item("ConvertCallback")

        If web_Converter.Exists("Instance") Then
            Dim web_Instance As Object
            Set web_Instance = web_Converter.Item("Instance")
            ConvertToFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Obj)
        Else
            ConvertToFormat = Application.Run(web_Callback, Obj)
        End If
#Else
    LogWarning "Custom formatting is disabled. To use WebFormat.Custom, enable custom formatting with the EnableCustomFormatting flag in WebHelpers"
#End If
    Case Else
        If VBA.VarType(Obj) = vbString Then
            ConvertToFormat = Obj
        End If
    End Select
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred during conversion" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    LogError web_ErrorDescription, "WebHelpers.ConvertToFormat", 11001
    Err.Raise 11001, "WebHelpers.ConvertToFormat", web_ErrorDescription
End Function

Public Function JoinUrl(LeftSide As String, RightSide As String) As String
    If Left(RightSide, 1) = "/" Then
        RightSide = Right(RightSide, Len(RightSide) - 1)
    End If
    If Right(LeftSide, 1) = "/" Then
        LeftSide = Left(LeftSide, Len(LeftSide) - 1)
    End If

    If LeftSide <> "" And RightSide <> "" Then
        JoinUrl = LeftSide & "/" & RightSide
    Else
        JoinUrl = LeftSide & RightSide
    End If
End Function

Public Function UrlEncode(Text As Variant, _
    Optional SpaceAsPlus As Boolean = False, Optional EncodeUnsafe As Boolean = True, _
    Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.StrictUrlEncoding) As String

    If SpaceAsPlus = True Then
        LogWarning "SpaceAsPlus is deprecated and will be removed in VBA-Web v5. " & _
            "Use EncodingMode:=FormUrlEncoding instead", "WebHelpers.UrlEncode"
    End If
    If EncodeUnsafe = False Then
        LogWarning "EncodeUnsafe has been removed as it was based on an outdated url encoding specification. " & _
            "Use EncodingMode:=CookieUrlEncoding to approximate this behavior", "WebHelpers.UrlEncode"
    End If

    Dim web_UrlVal As String
    Dim web_StringLen As Long

    web_UrlVal = VBA.CStr(Text)
    web_StringLen = VBA.Len(web_UrlVal)

    If web_StringLen > 0 Then
        Dim web_Result() As String
        Dim web_i As Long
        Dim web_CharCode As Integer
        Dim web_Char As String
        Dim web_Space As String
        ReDim web_Result(web_StringLen)


        If SpaceAsPlus Or EncodingMode = UrlEncodingMode.FormUrlEncoding Then
            web_Space = "+"
        Else
            web_Space = "%20"
        End If

        For web_i = 1 To web_StringLen
            web_Char = VBA.Mid$(web_UrlVal, web_i, 1)
            web_CharCode = VBA.Asc(web_Char)

            Select Case web_CharCode
                Case 65 To 90, 97 To 122
                    web_Result(web_i) = web_Char
                Case 48 To 57
                    web_Result(web_i) = web_Char
                Case 45, 46, 95
                    web_Result(web_i) = web_Char

                Case 32
                    web_Result(web_i) = web_Space

                Case 33, 36, 38, 39, 40, 41, 43, 58, 61, 64
                    If EncodingMode = UrlEncodingMode.PathUrlEncoding Or EncodingMode = UrlEncodingMode.CookieUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 35, 45, 46, 47, 60, 62, 63, 91, 93, 94, 95, 96, 123, 124, 125
                    If EncodingMode = UrlEncodingMode.CookieUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 42
                    If EncodingMode = UrlEncodingMode.FormUrlEncoding _
                        Or EncodingMode = UrlEncodingMode.PathUrlEncoding _
                        Or EncodingMode = UrlEncodingMode.CookieUrlEncoding Then

                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 44, 59
                    If EncodingMode = UrlEncodingMode.PathUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 126
                    If EncodingMode = UrlEncodingMode.FormUrlEncoding Or EncodingMode = UrlEncodingMode.QueryUrlEncoding Then
                        web_Result(web_i) = "%7E"
                    Else
                        web_Result(web_i) = web_Char
                    End If

                Case 0 To 15
                    web_Result(web_i) = "%0" & VBA.Hex(web_CharCode)
                Case Else
                    web_Result(web_i) = "%" & VBA.Hex(web_CharCode)

            End Select
        Next web_i
        UrlEncode = VBA.Join$(web_Result, "")
    End If
End Function

Public Function ConvertToUrlEncoded(Obj As Variant, Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.FormUrlEncoding) As String
    Dim web_Encoded As String

    If TypeOf Obj Is Collection Then
        Dim web_KeyValue As Dictionary

        For Each web_KeyValue In Obj
            If VBA.Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_KeyValue.Item("Key"), web_KeyValue.Item("Value"), EncodingMode)
        Next web_KeyValue
    Else
        Dim web_Key As Variant

        For Each web_Key In Obj.Keys()
            If Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_Key, Obj(web_Key), EncodingMode)
        Next web_Key
    End If

    ConvertToUrlEncoded = web_Encoded
End Function

Public Function ParseByFormat(Value As String, Format As WebFormat, _
    Optional CustomFormat As String = "", Optional Bytes As Variant) As Object

    On Error GoTo web_ErrorHandling

    If Value = "" And CustomFormat = "" Then
        Exit Function
    End If

    Select Case Format
    Case WebFormat.Json
        Set ParseByFormat = ParseJson(Value)
    Case WebFormat.FormUrlEncoded
        Set ParseByFormat = ParseUrlEncoded(Value)
    Case WebFormat.Xml
        Set ParseByFormat = ParseXml(Value)
    Case WebFormat.Custom
#If EnableCustomFormatting Then
        Dim web_Converter As Dictionary
        Dim web_Callback As String

        Set web_Converter = web_GetConverter(CustomFormat)
        web_Callback = web_Converter.Item("ParseCallback")

        If web_Converter.Exists("Instance") Then
            Dim web_Instance As Object
            Set web_Instance = web_Converter.Item("Instance")

            If web_Converter.Item("ParseType") = "Binary" Then
                Set ParseByFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Bytes)
            Else
                Set ParseByFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Value)
            End If
        Else
            If web_Converter.Item("ParseType") = "Binary" Then
                Set ParseByFormat = Application.Run(web_Callback, Bytes)
            Else
                Set ParseByFormat = Application.Run(web_Callback, Value)
            End If
        End If
#End If
    End Select
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred during parsing" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    Err.Raise 11000, "WebHelpers.ParseByFormat", web_ErrorDescription
End Function

Public Function ParseJson(ByVal JsonString As String) As Object
    Dim json_Index As Long
    json_Index = 1

    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

    json_SkipSpaces JsonString, json_Index
    Select Case VBA.Mid$(JsonString, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(JsonString, json_Index)
    Case "["
        Set ParseJson = json_ParseArray(JsonString, json_Index)
    Case Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
    End Select
End Function

Public Function ParseUrlEncoded(Encoded As String) As Dictionary
    Dim web_Items As Variant
    Dim web_i As Integer
    Dim web_Parts As Variant
    Dim web_Key As String
    Dim web_Value As Variant
    Dim web_Parsed As New Dictionary

    web_Items = VBA.Split(Encoded, "&")
    For web_i = LBound(web_Items) To UBound(web_Items)
        web_Parts = VBA.Split(web_Items(web_i), "=")

        If UBound(web_Parts) - LBound(web_Parts) >= 1 Then
            web_Key = UrlDecode(VBA.CStr(web_Parts(LBound(web_Parts))))
            web_Value = UrlDecode(VBA.CStr(web_Parts(LBound(web_Parts) + 1)))

            web_Parsed.Item(web_Key) = web_Value
        End If
    Next web_i

    Set ParseUrlEncoded = web_Parsed
End Function

Public Function ParseXml(Encoded As String) As Object
    Dim web_ErrorMsg As String

    web_ErrorMsg = "XML is not currently supported (An updated parser is being created that supports Mac and Windows)." & vbNewLine & _
        "To use XML parsing for Windows currently, use the instructions found here:" & vbNewLine & _
        vbNewLine & _
        "https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0"

    LogError web_ErrorMsg, "WebHelpers.ParseXml", 11099
    Err.Raise 11099, "WebHeleprs.ParseXml", web_ErrorMsg
End Function

Public Function UrlDecode(Encoded As String, _
    Optional PlusAsSpace As Boolean = True, _
    Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.StrictUrlEncoding) As String

    Dim web_StringLen As Long
    web_StringLen = VBA.Len(Encoded)

    If web_StringLen > 0 Then
        Dim web_i As Long
        Dim web_Result As String
        Dim web_Temp As String

        For web_i = 1 To web_StringLen
            web_Temp = VBA.Mid$(Encoded, web_i, 1)

            If web_Temp = "+" And _
                (PlusAsSpace _
                 Or EncodingMode = UrlEncodingMode.FormUrlEncoding _
                 Or EncodingMode = UrlEncodingMode.QueryUrlEncoding) Then

                web_Temp = " "
            ElseIf web_Temp = "%" And web_StringLen >= web_i + 2 Then
                web_Temp = VBA.Mid$(Encoded, web_i + 1, 2)
                web_Temp = VBA.Chr(VBA.CInt("&H" & web_Temp))

                web_i = web_i + 2
            End If


            web_Result = web_Result & web_Temp
        Next web_i

        UrlDecode = web_Result
    End If
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary
    Dim json_Key As String
    Dim json_NextChar As String

    Set json_ParseObject = New Dictionary
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            Else
                json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            End If
        Loop
    End If
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_Char As String
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_ParseKey = json_ParseKey & json_Char
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
    End If

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    json_SkipSpaces json_String, json_Index

    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        Select Case json_Char
        Case "\"
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)

            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_Buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function

Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
    Dim json_Char As String
    Dim json_Value As String
    Dim json_IsLargeNumber As Boolean

    json_SkipSpaces json_String, json_Index

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        If VBA.InStr("+-0123456789.eE", json_Char) Then
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                json_ParseNumber = json_Value
            Else
                json_ParseNumber = VBA.Val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
    Set json_ParseArray = New Collection

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_ParseArray.Add json_ParseValue(json_String, json_Index)
        Loop
    End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)

    Dim json_StartIndex As Long
    Dim json_StopIndex As Long

    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub json_BufferAppend(ByRef json_Buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long

    json_AppendLength = VBA.Len(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        Dim json_AddedLength As Long
        json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)

        json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
        json_BufferLength = json_BufferLength + json_AddedLength
    End If

    Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
    json_BufferPosition = json_BufferPosition + json_AppendLength
End Sub

Public Function ExecuteInShell(web_Command As String) As ShellResult
#If Mac Then
#If VBA7 Then
    Dim web_File As LongPtr
#Else
    Dim web_File As Long
#End If

    Dim web_Chunk As String
    Dim web_Read As Long

    On Error GoTo web_Cleanup

    web_File = web_popen(web_Command, "r")

    If web_File = 0 Then
        Exit Function
    End If

    Do While web_feof(web_File) = 0
        web_Chunk = VBA.Space$(50)
        web_Read = CLng(web_fread(web_Chunk, 1, Len(web_Chunk) - 1, web_File))
        If web_Read > 0 Then
            web_Chunk = VBA.Left$(web_Chunk, web_Read)
            ExecuteInShell.Output = ExecuteInShell.Output & web_Chunk
        End If
        
        VBA.DoEvents
    Loop

web_Cleanup:

    ExecuteInShell.ExitCode = CLng(web_pclose(web_File))
#End If
End Function

Public Function GetUrlParts(Url As String) As Dictionary
    Dim web_Parts As New Dictionary

    On Error GoTo web_ErrorHandling

#If Mac Then

    Dim web_AddedProtocol As Boolean
    Dim web_Command As String
    Dim web_Results As Variant
    Dim web_ResultPart As Variant
    Dim web_EqualsIndex As Long
    Dim web_Key As String
    Dim web_Value As String

    If InStr(1, Url, "://") <= 0 Then
        web_AddedProtocol = True
        If InStr(1, Url, "//") = 1 Then
            Url = "http" & Url
        Else
            Url = "http://" & Url
        End If
    End If

    web_Command = "perl -e '{use URI::URL;" & vbNewLine & _
        "$url = new URI::URL """ & Url & """;" & vbNewLine & _
        "print ""Protocol="" . $url->scheme;" & vbNewLine & _
        "print "" | Host="" . $url->host;" & vbNewLine & _
        "print "" | Port="" . $url->port;" & vbNewLine & _
        "print "" | FullPath="" . $url->full_path;" & vbNewLine & _
        "print "" | Hash="" . $url->frag;" & vbNewLine & _
    "}'"

    web_Results = Split(ExecuteInShell(web_Command).Output, " | ")
    For Each web_ResultPart In web_Results
        web_EqualsIndex = InStr(1, web_ResultPart, "=")
        web_Key = Trim(VBA.Mid$(web_ResultPart, 1, web_EqualsIndex - 1))
        web_Value = Trim(VBA.Mid$(web_ResultPart, web_EqualsIndex + 1))

        If web_Key = "FullPath" Then
            Dim QueryIndex As Integer

            QueryIndex = InStr(1, web_Value, "?")
            If QueryIndex > 0 Then
                web_Parts.Add "Path", Mid$(web_Value, 1, QueryIndex - 1)
                web_Parts.Add "Querystring", Mid$(web_Value, QueryIndex + 1)
            Else
                web_Parts.Add "Path", web_Value
                web_Parts.Add "Querystring", ""
            End If
        Else
            web_Parts.Add web_Key, web_Value
        End If
    Next web_ResultPart

    If web_AddedProtocol And web_Parts.Exists("Protocol") Then
        web_Parts.Item("Protocol") = ""
    End If
#Else
    If web_pDocumentHelper Is Nothing Or web_pElHelper Is Nothing Then
        Set web_pDocumentHelper = CreateObject("htmlfile")
        Set web_pElHelper = web_pDocumentHelper.createElement("a")
    End If

    web_pElHelper.href = Url
    web_Parts.Add "Protocol", Replace(web_pElHelper.Protocol, ":", "", Count:=1)
    web_Parts.Add "Host", web_pElHelper.hostname
    web_Parts.Add "Port", web_pElHelper.port
    web_Parts.Add "Path", web_pElHelper.pathname
    web_Parts.Add "Querystring", Replace(web_pElHelper.Search, "?", "", Count:=1)
    web_Parts.Add "Hash", Replace(web_pElHelper.Hash, "#", "", Count:=1)
#End If

    If web_Parts.Item("Protocol") = "localhost" Then
        Dim PathParts As Variant
        PathParts = Split(web_Parts("Path"), "/")

        web_Parts.Item("Port") = PathParts(0)
        web_Parts.Item("Protocol") = ""
        web_Parts.Item("Host") = "localhost"
        web_Parts.Item("Path") = Replace(web_Parts.Item("Path"), web_Parts.Item("Port"), "", Count:=1)
    End If
    If Left(web_Parts.Item("Path"), 1) <> "/" Then
        web_Parts.Item("Path") = "/" & web_Parts.Item("Path")
    End If

    Set GetUrlParts = web_Parts
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while getting url parts" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    LogError web_ErrorDescription, "WebHelpers.GetUrlParts", 11003
    Err.Raise 11003, "WebHelpers.GetUrlParts", web_ErrorDescription
End Function

Public Function PrepareTextForPrintf(ByVal web_Text As String) As String
    web_Text = VBA.Replace(web_Text, "\", "\\")
    web_Text = VBA.Replace(web_Text, "`", "\`")
    web_Text = VBA.Replace(web_Text, "$", "\$")
    web_Text = VBA.Replace(web_Text, "%", "%%")
    web_Text = VBA.Replace(web_Text, """", "\""")

    web_Text = """" & web_Text & """"

    web_Text = VBA.Replace(web_Text, "!", """'!'""")

    If VBA.Left$(web_Text, 3) = """""'" Then
        web_Text = VBA.Right$(web_Text, VBA.Len(web_Text) - 2)
    End If
    If VBA.Right$(web_Text, 3) = "'""""" Then
        web_Text = VBA.Left$(web_Text, VBA.Len(web_Text) - 2)
    End If

    PrepareTextForPrintf = web_Text
End Function

Public Function PrepareTextForShell(ByVal web_Text As String) As String
    web_Text = VBA.Replace(web_Text, "\", "\\")
    web_Text = VBA.Replace(web_Text, "`", "\`")
    web_Text = VBA.Replace(web_Text, "$", "\$")
    web_Text = VBA.Replace(web_Text, "%", "\%")
    web_Text = VBA.Replace(web_Text, """", "\""")

    web_Text = """" & web_Text & """"

    web_Text = VBA.Replace(web_Text, "!", """'!'""")

    If VBA.Left$(web_Text, 3) = """""'" Then
        web_Text = VBA.Right$(web_Text, VBA.Len(web_Text) - 2)
    End If
    If VBA.Right$(web_Text, 3) = "'""""" Then
        web_Text = VBA.Left$(web_Text, VBA.Len(web_Text) - 2)
    End If

    PrepareTextForShell = web_Text
End Function

Public Function StringToAnsiBytes(web_Text As String) As Byte()
    Dim web_Bytes() As Byte
    Dim web_AnsiBytes() As Byte
    Dim web_ByteIndex As Long
    Dim web_AnsiIndex As Long

    If VBA.Len(web_Text) > 0 Then
        web_Bytes = web_Text
        ReDim web_AnsiBytes(VBA.Int(UBound(web_Bytes) / 2))

        web_AnsiIndex = LBound(web_Bytes)
        For web_ByteIndex = LBound(web_Bytes) To UBound(web_Bytes) Step 2
            web_AnsiBytes(web_AnsiIndex) = web_Bytes(web_ByteIndex)
            web_AnsiIndex = web_AnsiIndex + 1
        Next web_ByteIndex
    End If

    StringToAnsiBytes = web_AnsiBytes
End Function
