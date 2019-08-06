Attribute VB_Name = "ABRA_API"
Option Explicit

Public errorInRequest As Boolean
Public listCalculating As Boolean

Private Function SendRequest(Url As String, Username As String, Password As String) As String
 If errorInRequest And listCalculating Then
  SendRequest = -1
  Exit Function
 End If
 On Error GoTo eh
  Dim timeout As Long
  timeout = Sheets("API").Range("Timeout").Value * 1000
  
  Dim Client As New WebClient
  Dim Request As New WebRequest
  Dim Response As WebResponse
  Dim Auth As New HttpBasicAuthenticator
 
  Client.BaseUrl = Url
  Client.TimeoutMs = timeout

  Auth.Setup Username, Password
  Set Client.Authenticator = Auth
 
  Set Response = Client.Execute(Request)

  If Response.StatusCode = WebStatusCode.Ok Then
   SendRequest = Response.Content
  Else
   showError (Response.StatusCode & ": " & Response.StatusDescription & vbCrLf & Response.Content)
  End If
Done:
  Exit Function
eh:
  showError (Err.Description)
End Function

Private Sub showError(Message As String)
  errorInRequest = True
  MsgBox Message
End Sub

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
  Rem Funkce pro konverzi œ‹tó z form‡tu "343A,524B..." na "343MD,524D..."
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
  AbraTurnover = GetTurnover(Url, Username, Password, Accounts, True, DateFrom, DateTo, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
End Function

Function GetTurnoverSimple(Url As String, Username As String, Password As String, DateFrom As Date, DateTo As Date, Conditions As String) As Double
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
  AbraBalance = GetBalance(Url, Username, Password, Accounts, True, DateTo, Divisions, False)
End Function

Function GetSale(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False, Optional BusOrders As String = "", Optional BusOrdersWithChildren As Boolean = False, Optional BusTransactions As String = "", Optional BusTransactionsWithChildren As Boolean = False, Optional BusProjects As String = "", Optional BusProjectsWithChildren As Boolean = False, Optional Firms As String = "") As Double
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
  AbraSale = GetSale(Url, Username, Password, InformationType, DateFrom, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
End Function

Function GetReceivable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = 0, Optional DocDateTo As Date = 0, Optional DueDateFrom As Date = 0, Optional DueDateTo As Date = 0, Optional Firms As String = "", Optional ACurrency As String = "") As Double
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
  AbraReceivable = GetReceivable(Url, Username, Password, InformationType, DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency)
End Function

Function GetPayable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = 0, Optional DocDateTo As Date = 0, Optional DueDateFrom As Date = 0, Optional DueDateTo As Date = 0, Optional Firms As String = "", Optional ACurrency As String = "") As Double
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
  AbraPayable = GetPayable(Url, Username, Password, InformationType, DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency)
End Function

Function GetStock(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "") As Double
  Dim mURL As String
  mURL = Url & "/utils/stock"
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
  Dim mResult As String
  mResult = SendRequest(mURL, Username, Password)
  GetStock = Val(mResult)
End Function

Function AbraStock(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "") As Double
  AbraStock = GetStock(Url, Username, Password, InformationType, DateFrom, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores)
End Function

Function GetMoves(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False, Optional BusOrders As String = "", Optional BusOrdersWithChildren As Boolean = False, Optional BusTransactions As String = "", Optional BusTransactionsWithChildren As Boolean = False, Optional BusProjects As String = "", Optional BusProjectsWithChildren As Boolean = False, Optional Firms As String = "") As Double
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
  AbraMoves = GetMoves(Url, Username, Password, InformationType, DateFrom, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
End Function

Function GetDepreciation(Url As String, Username As String, Password As String, InformationType As String, Optional DateFrom As Date = 0, Optional DateTo As Date = 0, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional AssetLocationsWithChildren As Boolean = False, Optional Responsibles As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False) As Double
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

Function AbraDepreciation(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional Responsibles As String = "", Optional Divisions As String = "") As Double
  AbraDepreciation = GetDepreciation(Url, Username, Password, InformationType, DateFrom, DateTo, AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, False, Responsibles, Divisions, False)
End Function

Function GetAsset(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional AssetLocationsWithChildren As Boolean = False, Optional Responsibles As String = "", Optional EvidenceDivisions As String = "", Optional EvidenceDivisionsWithChildren As Boolean = False) As Double
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
  AbraAsset = GetAsset(Url, Username, Password, InformationType, DateTo, AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, False, Responsibles, EvidenceDivisions, False)
End Function

Function GetPayroll(Url As String, Username As String, Password As String, InformationType As String, WagePeriods As String, Optional EmployPatterns As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Boolean = False) As Double
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
  AbraPayroll = GetPayroll(Url, Username, Password, InformationType, WagePeriods, EmployPatterns, Divisions, False)
End Function


