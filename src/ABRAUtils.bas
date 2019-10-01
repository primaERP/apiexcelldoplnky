Attribute VB_Name = "ABRAUtils"
Option Explicit

Public StopCalculating As Boolean
Public LastTimeError As Double

Private Function SendRequest(Url As String, Username As String, Password As String) As String
    Dim Timeout As Long
    Dim ConnectionTimeoutMs As Long
    Dim ConnectionTimeoutSec As Long
    
    Timeout = Sheets("API").Range("Timeout").Value * 1000
    ConnectionTimeoutSec = Sheets("API").Range("ConnectionTimeout").Value
    ConnectionTimeoutMs = ConnectionTimeoutSec * 1000
    
    If StopCalculating Then
        Dim NowTime As Double
        NowTime = TimeNow()
        If NowTime > (LastTimeError + ConnectionTimeoutSec + 5) Then
            StopCalculating = False
        End If
    End If
    
    If StopCalculating Then
        SendRequest = -1
        Exit Function
    End If
    
    On Error GoTo eh
  
    Dim Client As New WebClient
    Dim Request As New WebRequest
    Dim Response As WebResponse
    Dim Auth As New HttpBasicAuthenticator
         
    Client.BaseUrl = Url
    Client.TimeoutMs = Timeout
    Client.ConnectionTimeoutMs = ConnectionTimeoutMs
        
    Auth.Setup Username, Password
    Set Client.Authenticator = Auth
         
    Set Response = Client.Execute(Request)
        
    If Response.StatusCode = WebStatusCode.Ok Then
        SendRequest = Response.Content
    Else
        #If Mac Then
            ProcessError Response.StatusCode & ": " & Response.StatusDescription & vbCrLf & UTF8_Decode(Response.Content), ConnectionTimeoutSec
        #Else
            Dim responseText As String
            responseText = AlterCharset(Response.Content, "latin1", "UTF-8")
            ProcessError Response.StatusCode & ": " & Response.StatusDescription & vbCrLf & responseText, ConnectionTimeoutSec
        #End If
    End If
Done:
    Exit Function
eh:
    ProcessError Err.Description, ConnectionTimeoutSec
End Function

' For Mac
Function UTF8_Decode(ByVal sStr As String)
    Dim l As Long, sUTF8 As String, iChar As Integer, iChar2 As Integer
    For l = 1 To Len(sStr)
        iChar = Asc(Mid(sStr, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
            iChar2 = Asc(Mid(sStr, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(Mid(sStr, l + 1, 1))
            iChar3 = Asc(Mid(sStr, l + 2, 1))
            sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
            l = l + 2
        End If
            Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    UTF8_Decode = sUTF8
End Function


' This will alter charset of a string from 1-byte charset
' to another 1-byte charset
Function AlterCharset(Str As String, FromCharset As String, ToCharset As String)
  Dim Bytes
  If Str <> "" Then
    Bytes = StringToBytes(Str, FromCharset)
    AlterCharset = BytesToString(Bytes, ToCharset)
  Else
    AlterCharset = ""
  End If
End Function

' accept a string and convert it to Bytes array in the selected Charset
Function StringToBytes(Str, Charset)
  Dim Stream: Set Stream = CreateObject("ADODB.Stream")
  Stream.Type = 2
  Stream.Charset = Charset
  Stream.Open
  Stream.WriteText Str
  Stream.Flush
  Stream.Position = 0
  ' rewind stream and read Bytes
  Stream.Type = 1
  StringToBytes = Stream.Read
  Stream.Close
  Set Stream = Nothing
End Function

' accept Bytes array and convert it to a string using the selected charset
Function BytesToString(Bytes, Charset)
  Dim Stream: Set Stream = CreateObject("ADODB.Stream")
  Stream.Charset = Charset
  Stream.Type = 1
  Stream.Open
  Stream.Write Bytes
  Stream.Flush
  Stream.Position = 0
  ' rewind stream and read text
  Stream.Type = 2
  BytesToString = Stream.ReadText
  Stream.Close
  Set Stream = Nothing
End Function

Private Sub ProcessError(ErrorDescription As String, ConnectionTimeoutSec As Long)
    Dim NowTime As Double
    NowTime = TimeNow()
    
    If NowTime > (LastTimeError + ConnectionTimeoutSec + 5) Then
        ShowError (ErrorDescription)
        LastTimeError = NowTime
    Else
        StopCalculating = True
    End If
End Sub

Private Sub ShowError(Message As String)
  MsgBox Message
End Sub

Function TimeNow() As Double
    'TimeNow = ((Now - 25569) * 86400000) - 3600000
    TimeNow = (Now - Date) * 24 * 60 * 60
End Function

Private Function StrPadZero2(Value As String) As String
  If Len(Value) = 1 Then
    StrPadZero2 = "0" + Value
  Else
    StrPadZero2 = Value
  End If
End Function

Private Function DateToISO8601(Value As Date) As String
  Dim d As String
  Dim m As String
  Dim y As String
  d = StrPadZero2(Day(Value))
  m = StrPadZero2(Month(Value))
  y = Year(Value)
  DateToISO8601 = y & "-" & m & "-" & d
End Function

Private Function CorrectAccounts(AAccount As String) As String
  Rem Funkce pro konverzi uctu z formatu "343A,524B..." na "343MD,524D..."
  Dim mAccArray, mACC2
  mAccArray = Split(AAccount, ",")
  AAccount = ""
  For Each mACC2 In mAccArray
    If Right(mACC2, 1) = "A" Then
      mACC2 = Left(mACC2, Len(mACC2) - 1) & "MD"
    End If
    If Right(mACC2, 1) = "B" Then
      mACC2 = Left(mACC2, Len(mACC2) - 1) & "D"
    End If
    If AAccount <> "" Then AAccount = AAccount & ","
    AAccount = AAccount & mACC2
  Next
  CorrectAccounts = AAccount
End Function
                    
Function GetTurnover(Url As String, Username As String, Password As String, Accounts As String, IncludeRequests As Boolean, DateFrom As Date, DateTo As Date, Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False, Optional BusOrders As String = "", Optional BusOrdersWithChildren As Boolean = False, Optional BusTransactions As String = "", Optional BusTransactionsWithChildren As Boolean = False, Optional BusProjects As String = "", Optional BusProjectsWithChildren As Boolean = False, Optional Firms As String = "") As Double
  Attribute GetTurnover.VB_Description = "Slouží k získávání informací o obratech na úctech.\n Povinné parametry: Url, Username, Password, Accounts, IncludeRequests, DateFrom, DateTo\n Nepovinné parametry: Divisions, DivisionsWithChildren, BusOrders, BusOrdersWithChildren, BusTransactions, BusTransactionsWithChildren, BusProjects, BusProjectsWithChildren, Firms."
  Dim mURL As String
  mURL = Url & "/bookentries/turnover"
  mURL = mURL & "?date-from=" & DateToISO8601(DateFrom)
  mURL = mURL & "&date-to=" & DateToISO8601(DateTo)
  mURL = mURL & "&accounts=" & CorrectAccounts(Accounts)
  If IncludeRequests Then
    mURL = mURL & "&include-requests=true"
  End If
  If Divisions <> "" Then
    mURL = mURL & "&divisions=" & Divisions
    If DivisionsWithChildren Then
      mURL = mURL & "&divisions-with-children=true"
    End If
  End If
  If BusOrders <> "" Then
    mURL = mURL & "&busorders=" & BusOrders
    If BusOrdersWithChildren Then
      mURL = mURL & "&busorders-with-children=true"
    End If
  End If
  If BusTransactions <> "" Then
    mURL = mURL & "&bustransactions=" & BusTransactions
    If BusTransactionsWithChildren Then
      mURL = mURL & "&bustransactions-with-children=true"
    End If
  End If
  If BusProjects <> "" Then
    mURL = mURL & "&busprojects=" & BusProjects
    If BusProjectsWithChildren Then
      mURL = mURL & "&busprojects-with-children=true"
    End If
  End If
  If Firms <> "" Then
    mURL = mURL & "&firms=" & Firms
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetTurnover = Val(mResult)
End Function

Function AbraTurnover(Url As String, Username As String, Password As String, Accounts As String, DateFrom As Date, DateTo As Date, Optional Divisions As String = "", Optional BusOrders As String = "", Optional BusTransactions As String = "", Optional BusProjects As String = "", Optional Firms As String = "") As Double
  Attribute AbraTurnover.VB_Description = "Slouží k získávání informací o obratech na úctech.\n Povinné parametry: Url, Username, Password, Accounts, DateFrom, DateTo\n Nepovinné parametry: Divisions, BusOrders, BusTransactions, BusProjects, Firms."
  AbraTurnover = GetTurnover(Url, Username, Password, Accounts, True, DateFrom, DateTo, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
End Function

Function GetTurnoverSimple(Url As String, Username As String, Password As String, DateFrom As Date, DateTo As Date, Conditions As String) As Double
  Attribute GetTurnoverSimple.VB_Description = "Slouží k získávání informací o obratech na úctech.\n Povinné parametry: Url, Username, Password, DateFrom, DateTo, Conditions"
  Dim mURL As String
  mURL = Url & "/bookentries/turnover-simple"
  mURL = mURL & "?date-from=" & DateToISO8601(DateFrom)
  mURL = mURL & "&date-to=" & DateToISO8601(DateTo)
  mURL = mURL & "&conditions=" & Conditions
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetTurnoverSimple = Val(mResult)
End Function

Function GetBalance(Url As String, Username As String, Password As String, Accounts As String, IncludeRequests As Boolean, DateTo As Date, Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False) As Double
  Attribute GetBalance.VB_Description = "Slouží k získávání informací o zustatcích úctu k urcitému datu.\n Povinné parametry: Url, Username, Password, Accounts, IncludeRequests, DateTo\n Nepovinné parametry: Divisions, DivisionsWithChildren."
  Dim mURL As String
  mURL = Url & "/utils/balance"
  mURL = mURL & "?date-to=" & DateToISO8601(DateTo)
  mURL = mURL & "&accounts=" & CorrectAccounts(Accounts)
  If IncludeRequests Then
    mURL = mURL & "&include-requests=true"
  End If
  If Divisions <> "" Then
    mURL = mURL & "&divisions=" & Divisions
    If DivisionsWithChildren Then
      mURL = mURL & "&divisions-with-children=true"
    End If
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetBalance = Val(mResult)
End Function

Function AbraBalance(Url As String, Username As String, Password As String, Accounts As String, DateTo As Date, Optional Divisions As String = "") As Double
  Attribute AbraBalance.VB_Description = "Slouží k získávání informací o zustatcích úctu k urcitému datu.\n Povinné parametry: Url, Username, Password, Accounts, DateTo\n Nepovinný parametr: Divisions."
  AbraBalance = GetBalance(Url, Username, Password, Accounts, True, DateTo, Divisions, False)
End Function

Function GetSale(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False, Optional BusOrders As String = "", Optional BusOrdersWithChildren As Boolean = False, Optional BusTransactions As String = "", Optional BusTransactionsWithChildren As Boolean = False, Optional BusProjects As String = "", Optional BusProjectsWithChildren As Boolean = False, Optional Firms As String = "") As Double
  Attribute GetSale.VB_Description = "Slouží k získávání informací o prodejích skladových položek.\n Povinné parametry: Url, Username, Password, InformationType, DateFrom, DateTo\n Nepovinné parametry: StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, DivisionsWithChildren, BusOrders, BusOrdersWithChildren, BusTransactions, BusTransactionsWithChildren, BusProjects, BusProjectsWithChildren, Firms."
  Dim mURL As String
  mURL = Url & "/utils/sale"
  mURL = mURL & "?date-from=" & DateToISO8601(DateFrom)
  mURL = mURL & "&date-to=" & DateToISO8601(DateTo)
  mURL = mURL & "&information-type=" & InformationType
  If StoreMenuItems <> "" Then
    mURL = mURL & "&store-menu-items=" & StoreMenuItems
  End If
  If StoreCardCategories <> "" Then
    mURL = mURL & "&store-card-categories=" & StoreCardCategories
  End If
  If StoreCards <> "" Then
    mURL = mURL & "&store-cards=" & StoreCards
  End If
  If Stores <> "" Then
    mURL = mURL & "&stores=" & Stores
  End If
  If Firms <> "" Then
    mURL = mURL & "&firms=" & Firms
  End If
  If Divisions <> "" Then
    mURL = mURL & "&divisions=" & Divisions
    If DivisionsWithChildren Then
      mURL = mURL & "&divisions-with-children=true"
    End If
  End If
  If BusOrders <> "" Then
    mURL = mURL & "&busorders=" & BusOrders
    If BusOrdersWithChildren Then
      mURL = mURL & "busorders-with-children=true"
    End If
  End If
  If BusTransactions <> "" Then
    mURL = mURL & "&bustransactions=" & BusTransactions
    If BusTransactionsWithChildren Then
      mURL = mURL & "bustransactions-with-children=true"
    End If
  End If
  If BusProjects <> "" Then
    mURL = mURL & "&busprojects=" & BusProjects
    If BusProjectsWithChildren Then
      mURL = mURL & "busprojects-with-children=true"
    End If
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetSale = Val(mResult)
End Function

Function AbraSale(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional BusOrders As String = "", Optional BusTransactions As String = "", Optional BusProjects As String = "", Optional Firms As String = "") As Double
  Attribute GetSale.VB_Description = "Slouží k získávání informací o prodejích skladových položek.\n Povinné parametry: Url, Username, Password, InformationType, DateFrom, DateTo\n Nepovinné parametry: StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, BusOrders, BusTransactions, BusProjects, Firms."
  AbraSale = GetSale(Url, Username, Password, InformationType, DateFrom, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
End Function

Function GetReceivable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = 0, Optional DocDateTo As Date = 0, Optional DueDateFrom As Date = 0, Optional DueDateTo As Date = 0, Optional Firms As String = "", Optional ACurrency As String = "") As Double
  Attribute GetReceivable.VB_Description = "Slouží k získávání informací o pohledávkách.\n Povinné parametry: Url, Username, Password, InformationType\n Nepovinné parametry: DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency."
  Dim mURL As String
  mURL = Url & "/utils/receivable"
  mURL = mURL & "?information-type=" & InformationType
  If DocDateFrom <> 0 Then
    mURL = mURL & "&doc-date-from=" & DateToISO8601(DocDateFrom)
  End If
  If DocDateTo <> 0 Then
    mURL = mURL & "&doc-date-to=" & DateToISO8601(DocDateTo)
  End If
  If DueDateFrom <> 0 Then
    mURL = mURL & "&due-date-from=" & DateToISO8601(DueDateFrom)
  End If
  If DueDateTo <> 0 Then
    mURL = mURL & "&due-date-to=" & DateToISO8601(DueDateTo)
  End If
  If Firms <> "" Then
    mURL = mURL & "&firms=" & Firms
  End If
  If ACurrency <> "" Then
    mURL = mURL & "&currency=" & ACurrency
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetReceivable = Val(mResult)
End Function

Function AbraReceivable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = 0, Optional DocDateTo As Date = 0, Optional DueDateFrom As Date = 0, Optional DueDateTo As Date = 0, Optional Firms As String = "", Optional ACurrency As String = "") As Double
  Attribute AbraReceivable.VB_Description = "Slouží k získávání informací o pohledávkách.\n Povinné parametry: Url, Username, Password, InformationType\n Nepovinné parametry: DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency."
  AbraReceivable = GetReceivable(Url, Username, Password, InformationType, DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency)
End Function

Function GetPayable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = 0, Optional DocDateTo As Date = 0, Optional DueDateFrom As Date = 0, Optional DueDateTo As Date = 0, Optional Firms As String = "", Optional ACurrency As String = "") As Double
  Attribute GetPayable.VB_Description = "Slouží k získávání informací o závazcích.\n Povinné parametry: Url, Username, Password, InformationType\n Nepovinné parametry: DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency."
  Dim mURL As String
  mURL = Url & "/utils/payable"
  mURL = mURL & "?information-type=" & InformationType
  If DocDateFrom <> 0 Then
    mURL = mURL & "&doc-date-from=" & DateToISO8601(DocDateFrom)
  End If
  If DocDateTo <> 0 Then
    mURL = mURL & "&doc-date-to=" & DateToISO8601(DocDateTo)
  End If
  If DueDateFrom <> 0 Then
    mURL = mURL & "&due-date-from=" & DateToISO8601(DueDateFrom)
  End If
  If DueDateTo <> 0 Then
    mURL = mURL & "&due-date-to=" & DateToISO8601(DueDateTo)
  End If
  If Firms <> "" Then
    mURL = mURL & "&firms=" & Firms
  End If
  If ACurrency <> "" Then
    mURL = mURL & "&currency=" & ACurrency
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetPayable = Val(mResult)
End Function

Function AbraPayable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = 0, Optional DocDateTo As Date = 0, Optional DueDateFrom As Date = 0, Optional DueDateTo As Date = 0, Optional Firms As String = "", Optional ACurrency As String = "") As Double
  Attribute AbraPayable.VB_Description = "Slouží k získávání informací o závazcích.\n Povinné parametry: Url, Username, Password, InformationType\n Nepovinné parametry: DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency."
  AbraPayable = GetPayable(Url, Username, Password, InformationType, DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency)
End Function

Function GetStock(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "") As Double
  Attribute GetStock.VB_Description = "Slouží k získávání informací o stavu skladu.\n Povinné parametry: Url, Username, Password, InformationType, DateTo\n Nepovinné parametry: StoreMenuItems, StoreCardCategories, StoreCards, Stores."
  Dim mURL As String
  mURL = Url & "/utils/stock"
  mURL = mURL & "?date-to=" & DateToISO8601(DateTo)
  mURL = mURL & "&information-type=" & InformationType
  If StoreMenuItems <> "" Then
    mURL = mURL & "&store-menu-items=" & StoreMenuItems
  End If
  If StoreCardCategories <> "" Then
    mURL = mURL & "&store-card-categories=" & StoreCardCategories
  End If
  If StoreCards <> "" Then
    mURL = mURL & "&store-cards=" & StoreCards
  End If
  If Stores <> "" Then
    mURL = mURL & "&stores=" & Stores
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetStock = Val(mResult)
End Function

Function AbraStock(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "") As Double
  Attribute AbraStock.VB_Description = "Slouží k získávání informací o stavu skladu.\n Povinné parametry: Url, Username, Password, InformationType, DateTo\n Nepovinné parametry: StoreMenuItems, StoreCardCategories, StoreCards, Stores."
  AbraStock = GetStock(Url, Username, Password, InformationType, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores)
End Function

Function GetMoves(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False, Optional BusOrders As String = "", Optional BusOrdersWithChildren As Boolean = False, Optional BusTransactions As String = "", Optional BusTransactionsWithChildren As Boolean = False, Optional BusProjects As String = "", Optional BusProjectsWithChildren As Boolean = False, Optional Firms As String = "") As Double
  Attribute GetMoves.VB_Description = "Slouží k získávání informací o skladových pohybech.\n Povinné parametry: Url, Username, Password, InformationType, DateFrom, DateTo\n Nepovinné parametry: StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, DivisionsWithChildren, BusOrders, BusOrdersWithChildren, BusTransactions, BusTransactionsWithChildren, BusProjects, BusProjectsWithChildren, Firms."
  Dim mURL As String
  mURL = Url & "/utils/moves"
  mURL = mURL & "?date-from=" & DateToISO8601(DateFrom)
  mURL = mURL & "&date-to=" & DateToISO8601(DateTo)
  mURL = mURL & "&information-type=" & InformationType
  If StoreMenuItems <> "" Then
    mURL = mURL & "&store-menu-items=" & StoreMenuItems
  End If
  If StoreCardCategories <> "" Then
    mURL = mURL & "&store-card-categories=" & StoreCardCategories
  End If
  If StoreCards <> "" Then
    mURL = mURL & "&store-cards=" & StoreCards
  End If
  If Stores <> "" Then
    mURL = mURL & "&stores=" & Stores
  End If
  If Firms <> "" Then
    mURL = mURL & "&firms=" & Firms
  End If
  If Divisions <> "" Then
    mURL = mURL & "&divisions=" & Divisions
    If DivisionsWithChildren Then
      mURL = mURL & "&divisions-with-children=true"
    End If
  End If
  If BusOrders <> "" Then
    mURL = mURL & "&busorders=" & BusOrders
    If BusOrdersWithChildren Then
      mURL = mURL & "busorders-with-children=true"
    End If
  End If
  If BusTransactions <> "" Then
    mURL = mURL & "&bustransactions=" & BusTransactions
    If BusTransactionsWithChildren Then
      mURL = mURL & "bustransactions-with-children=true"
    End If
  End If
  If BusProjects <> "" Then
    mURL = mURL & "&busprojects=" & BusProjects
    If BusProjectsWithChildren Then
      mURL = mURL & "busprojects-with-children=true"
    End If
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetMoves = Val(mResult)
End Function

Function AbraMoves(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional BusOrders As String = "", Optional BusTransactions As String = "", Optional BusProjects As String = "", Optional Firms As String = "") As Double
  Attribute AbraMoves.VB_Description = "Slouží k získávání informací o skladových pohybech.\n Povinné parametry: Url, Username, Password, InformationType, DateFrom, DateTo\n Nepovinné parametry: StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, BusOrders, BusTransactions, BusProjects, Firms."
  AbraMoves = GetMoves(Url, Username, Password, InformationType, DateFrom, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
End Function

Function GetDepreciation(Url As String, Username As String, Password As String, InformationType As String, Optional DateFrom As Date = 0, Optional DateTo As Date = 0, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional AssetLocationsWithChildren As Boolean = False, Optional Responsibles As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False) As Double
  Attribute GetDepreciation.VB_Description = "Slouží k získávání informací o odpisech.\n Povinné parametry: Url, Username, Password, InformationType\n Nepovinné parametry: DateFrom, DateTo, AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, AssetLocationsWithChildren, Responsibles, Divisions, DivisionsWithChildren."
  Dim mURL As String
  mURL = Url & "/utils/depreciation"
  mURL = mURL & "?information-type=" & InformationType
  If DateFrom <> 0 Then
    mURL = mURL & "&date-from=" & DateToISO8601(DateFrom)
  End If
  If DateTo <> 0 Then
    mURL = mURL & "&date-to=" & DateToISO8601(DateTo)
  End If
  If AssetTypes <> "" Then
    mURL = mURL & "&asset-types=" & AssetTypes
  End If
  If AccDepreciationGroups <> "" Then
    mURL = mURL & "&acc-depreciation-groups=" & AccDepreciationGroups
  End If
  If TaxDepreciationGroups <> "" Then
    mURL = mURL & "&tax-depreciation-groups=" & TaxDepreciationGroups
  End If
  If Responsibles <> "" Then
    mURL = mURL & "&responsibles=" & Responsibles
  End If
  If AssetLocations <> "" Then
    mURL = mURL & "&asset-locations=" & AssetLocations
    If AssetLocationsWithChildren Then
      mURL = mURL & "&asset-locations-with-children=true"
    End If
  End If
  If Divisions <> "" Then
    mURL = mURL & "&divisions=" & Divisions
    If DivisionsWithChildren Then
      mURL = mURL & "&divisions-with-children=true"
    End If
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetDepreciation = Val(mResult)
End Function

Function AbraDepreciation(Url As String, Username As String, Password As String, InformationType As String, Optional DateFrom As Date = 0, Optional DateTo As Date = 0, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional Responsibles As String = "", Optional Divisions As String = "") As Double
  Attribute AbraDepreciation.VB_Description = "Slouží k získávání informací o odpisech.\n Povinné parametry: Url, Username, Password, InformationType\n Nepovinné parametry: DateFrom, DateTo, AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, Responsibles, Divisions."
  AbraDepreciation = GetDepreciation(Url, Username, Password, InformationType, DateFrom, DateTo, AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, False, Responsibles, Divisions, False)
End Function

Function GetAsset(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional AssetLocationsWithChildren As Boolean = False, Optional Responsibles As String = "", Optional EvidenceDivisions As String = "", Optional EvidenceDivisionsWithChildren As Boolean = False) As Double
  Attribute GetAsset.VB_Description = "Slouží k získávání informací o stavu majetku.\n Povinné parametry: Url, Username, Password, InformationType, DateTo\n Nepovinné parametry: AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, AssetLocationsWithChildren, Responsibles, EvidenceDivisions, EvidenceDivisionsWithChildren."
  Dim mURL As String
  mURL = Url & "/utils/asset"
  mURL = mURL & "?information-type=" & InformationType
  mURL = mURL & "&date-to=" & DateToISO8601(DateTo)
  If AssetTypes <> "" Then
    mURL = mURL & "&asset-types=" & AssetTypes
  End If
  If AccDepreciationGroups <> "" Then
    mURL = mURL & "&acc-depreciation-groups=" & AccDepreciationGroups
  End If
  If TaxDepreciationGroups <> "" Then
    mURL = mURL & "&tax-depreciation-groups=" & TaxDepreciationGroups
  End If
  If Responsibles <> "" Then
    mURL = mURL & "&responsibles=" & Responsibles
  End If
  If AssetLocations <> "" Then
    mURL = mURL & "&asset-locations=" & AssetLocations
    If AssetLocationsWithChildren Then
      mURL = mURL & "&asset-locations-with-children=true"
    End If
  End If
  If EvidenceDivisions <> "" Then
    mURL = mURL & "&evidence-divisions=" & EvidenceDivisions
    If EvidenceDivisionsWithChildren Then
      mURL = mURL & "&evidence-divisions-with-children=true"
    End If
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetAsset = Val(mResult)
End Function

Function AbraAsset(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional Responsibles As String = "", Optional EvidenceDivisions As String = "") As Double
  Attribute AbraAsset.VB_Description = "Slouží k získávání informací o stavu majetku.\n Povinné parametry: Url, Username, Password, InformationType, DateTo\n Nepovinné parametry: AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, Responsibles, EvidenceDivisions."
  AbraAsset = GetAsset(Url, Username, Password, InformationType, DateTo, AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, False, Responsibles, EvidenceDivisions, False)
End Function

Function GetPayroll(Url As String, Username As String, Password As String, InformationType As String, WagePeriods As String, Optional EmployPatterns As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False) As Double
  Attribute GetPayroll.VB_Description = "Slouží k získávání informací z oblasti mezd a personalistiky.\n Povinné parametry: Url, Username, Password, InformationType, WagePeriods\n Nepovinné parametry: EmployPatterns, Divisions, DivisionsWithChildren."
  Dim mURL As String
  mURL = Url & "/utils/payroll"
  mURL = mURL & "?information-type=" & InformationType
  mURL = mURL & "&wage-periods=" & WagePeriods
  If EmployPatterns <> "" Then
    mURL = mURL & "&employ-patterns=" & EmployPatterns
  End If
  If Divisions <> "" Then
    mURL = mURL & "&divisions=" & Divisions
    If DivisionsWithChildren Then
      mURL = mURL & "&divisions-with-children=true"
    End If
  End If
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetPayroll = Val(mResult)
End Function

Function AbraPayroll(Url As String, Username As String, Password As String, InformationType As String, WagePeriods As String, Optional EmployPatterns As String = "", Optional Divisions As String = "") As Double
  Attribute AbraPayroll.VB_Description = "Slouží k získávání informací z oblasti mezd a personalistiky.\n Povinné parametry: Url, Username, Password, InformationType, WagePeriods\n Nepovinné parametry: EmployPatterns, Divisions."
  AbraPayroll = GetPayroll(Url, Username, Password, InformationType, WagePeriods, EmployPatterns, Divisions, False)
End Function


