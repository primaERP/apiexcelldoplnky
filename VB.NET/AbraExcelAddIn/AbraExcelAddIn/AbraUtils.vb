Imports System.Net
Imports System.Net.Mime.MediaTypeNames
Imports System.Security.Cryptography
Imports System.Threading
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel
'Imports ExcelDna.Registration.VisualBasic


Public Module AbraUtils

    Private Function SendRequest(Url As String, Username As String, Password As String) As String
        Try
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim webClient As New System.Net.WebClient
            webClient.Headers.Add("Authorization", "Basic " + System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Username + ":" + Password)))
            Dim result As String = webClient.DownloadString(Url)
            Return result
        Catch e As Exception
            Return "0"
        End Try

    End Function
    Function convertBoolean(Optional param1 As Object = False) As Boolean
        convertBoolean = False
        If TypeOf param1 Is Boolean Then
            convertBoolean = param1
        ElseIf param1 <> Nothing Then
            Dim strParam = param1.ToString().ToLower.Trim


            If strParam = "0" Or strParam = "ne" Or strParam = "false" Then
                convertBoolean = False
            ElseIf strParam = "1" Or strParam = "ano" Or strParam = "true" Then
                convertBoolean = True
            End If
        End If
    End Function

    Private Function StrPadZero2(Value As String) As String
        If Len(Value) = 1 Then
            StrPadZero2 = "0" + Value
        Else
            StrPadZero2 = Value
        End If
    End Function

    Private Function CorrectAccounts(AAccount As String) As String
        REM Funkce pro konverzi uctu z formatu "343A,524B..." na "343MD,524D..."
        Dim mAccArray, mACC2
        mAccArray = Split(AAccount, ",")
        AAccount = ""
        For Each mACC2 In mAccArray
            If Right(mACC2, 1) = "A" Then
                mACC2 = Left(mACC2, Len(mACC2) - 1) & "MD"
            ElseIf Right(mACC2, 1) = "B" Then
                mACC2 = Left(mACC2, Len(mACC2) - 1) & "D"
            End If
            If AAccount <> "" Then AAccount = AAccount & ","
            AAccount = AAccount & mACC2
        Next
        CorrectAccounts = AAccount
    End Function

    Public Function EndsWith(str As String, ending As String) As Boolean
        Dim endingLen As Integer
        endingLen = Len(ending)
        EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
    End Function

    Private Function FilterBaseUrl(Url As String) As String
        If EndsWith(Url, "/") Then
            FilterBaseUrl = Left(Url, Len(Url) - 1)
        Else
            FilterBaseUrl = Url
        End If
    End Function

    Private Function DateToISO8601(dateparam As Date)
        DateToISO8601 = dateparam.Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.sssZ")
    End Function

    <ExcelFunction(Description:="Function for get turnover", IsThreadSafe:=True)>
    Function GetTurnover(Url As String, Username As String, Password As String, Accounts As String, IncludeRequests As Object, DateFrom As Date, DateTo As Date, Optional Divisions As String = "", Optional DivisionsWithChildren As Object = False, Optional BusOrders As String = "", Optional BusOrdersWithChildren As Object = False, Optional BusTransactions As String = "", Optional BusTransactionsWithChildren As Object = False, Optional BusProjects As String = "", Optional BusProjectsWithChildren As Object = False, Optional Firms As String = "") As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/bookentries/turnover"
        mURL = mURL & "?date-from=" & DateToISO8601(DateFrom)
        mURL = mURL & "&date-to=" & DateToISO8601(DateTo)
        mURL = mURL & "&accounts=" & CorrectAccounts(Accounts)
        If convertBoolean(IncludeRequests) Then
            mURL = mURL & "&include-requests=true"
        End If
        If Divisions <> "" Then
            mURL = mURL & "&divisions=" & Divisions
            If convertBoolean(DivisionsWithChildren) Then
                mURL = mURL & "&divisions-with-children=true"
            End If
        End If
        If BusOrders <> "" Then
            mURL = mURL & "&busorders=" & BusOrders
            If convertBoolean(BusOrdersWithChildren) Then
                mURL = mURL & "&busorders-with-children=true"
            End If
        End If
        If BusTransactions <> "" Then
            mURL = mURL & "&bustransactions=" & BusTransactions
            If convertBoolean(BusTransactionsWithChildren) Then
                mURL = mURL & "&bustransactions-with-children=true"
            End If
        End If
        If BusProjects <> "" Then
            mURL = mURL & "&busprojects=" & BusProjects
            If convertBoolean(BusProjectsWithChildren) Then
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

    <ExcelFunction(Description:="Function for get turnover", IsThreadSafe:=True)>
    Function AbraTurnover(Url As String, Username As String, Password As String, Accounts As String, DateFrom As Date, DateTo As Date, Optional Divisions As String = "", Optional BusOrders As String = "", Optional BusTransactions As String = "", Optional BusProjects As String = "", Optional Firms As String = "") As Double
        AbraTurnover = GetTurnover(Url, Username, Password, Accounts, True, DateFrom, DateTo, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
    End Function

    <ExcelFunction(Description:="Function for get turnover", IsThreadSafe:=True)>
    Function GetTurnoverSimple(Url As String, Username As String, Password As String, DateFrom As Date, DateTo As Date, Conditions As String) As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/bookentries/turnover-simple"
        mURL = mURL & "?date-from=" & DateToISO8601(DateFrom)
        mURL = mURL & "&date-to=" & DateToISO8601(DateTo)
        mURL = mURL & "&conditions=" & Conditions
        Dim mResult As String
        mResult = SendRequest(mURL, Username, Password)
        GetTurnoverSimple = Val(mResult)
    End Function

    <ExcelFunction(Description:="Function for get balance", IsThreadSafe:=True)>
    Function GetBalance(Url As String, Username As String, Password As String, Accounts As String, IncludeRequests As Object, DateTo As Date, Optional Divisions As String = "", Optional DivisionsWithChildren As Object = False) As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/balance"
        mURL = mURL & "?date-to=" & DateToISO8601(DateTo)
        mURL = mURL & "&accounts=" & CorrectAccounts(Accounts)
        If convertBoolean(IncludeRequests) Then
            mURL = mURL & "&include-requests=true"
        End If
        If Divisions <> "" Then
            mURL = mURL & "&divisions=" & Divisions
            If convertBoolean(DivisionsWithChildren) Then
                mURL = mURL & "&divisions-with-children=true"
            End If
        End If
        Dim mResult As String
        mResult = SendRequest(mURL, Username, Password)
        GetBalance = Val(mResult)
    End Function

    <ExcelFunction(Description:="Function for get balance", IsThreadSafe:=True)>
    Function AbraBalance(Url As String, Username As String, Password As String, Accounts As String, DateTo As Date, Optional Divisions As String = "") As Double
        AbraBalance = GetBalance(Url, Username, Password, Accounts, True, DateTo, Divisions, False)
    End Function

    <ExcelFunction(Description:="Function for get sale", IsThreadSafe:=True)>
    Function GetSale(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Object = False, Optional BusOrders As String = "", Optional BusOrdersWithChildren As Object = False, Optional BusTransactions As String = "", Optional BusTransactionsWithChildren As Object = False, Optional BusProjects As String = "", Optional BusProjectsWithChildren As Object = False, Optional Firms As String = "") As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/sale"
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
            If convertBoolean(DivisionsWithChildren) Then
                mURL = mURL & "&divisions-with-children=true"
            End If
        End If
        If BusOrders <> "" Then
            mURL = mURL & "&busorders=" & BusOrders
            If convertBoolean(BusOrdersWithChildren) Then
                mURL = mURL & "busorders-with-children=true"
            End If
        End If
        If BusTransactions <> "" Then
            mURL = mURL & "&bustransactions=" & BusTransactions
            If convertBoolean(BusTransactionsWithChildren) Then
                mURL = mURL & "bustransactions-with-children=true"
            End If
        End If
        If BusProjects <> "" Then
            mURL = mURL & "&busprojects=" & BusProjects
            If convertBoolean(BusProjectsWithChildren) Then
                mURL = mURL & "busprojects-with-children=true"
            End If
        End If
        Dim mResult As String
        mResult = SendRequest(mURL, Username, Password)
        GetSale = Val(mResult)
    End Function

    <ExcelFunction(Description:="Function for get sale", IsThreadSafe:=True)>
    Function AbraSale(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional BusOrders As String = "", Optional BusTransactions As String = "", Optional BusProjects As String = "", Optional Firms As String = "") As Double
        AbraSale = GetSale(Url, Username, Password, InformationType, DateFrom, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
    End Function

    <ExcelFunction(Description:="Function for get receivable", IsThreadSafe:=True)>
    Function GetReceivable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = Nothing, Optional DocDateTo As Date = Nothing, Optional DueDateFrom As Date = Nothing, Optional DueDateTo As Date = Nothing, Optional Firms As String = "", Optional ACurrency As String = "") As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/receivable"
        mURL = mURL & "?information-type=" & InformationType
        If DocDateFrom <> Nothing Then
            mURL = mURL & "&doc-date-from=" & DateToISO8601(DocDateFrom)
        End If
        If DocDateTo <> Nothing Then
            mURL = mURL & "&doc-date-to=" & DateToISO8601(DocDateTo)
        End If
        If DueDateFrom <> Nothing Then
            mURL = mURL & "&due-date-from=" & DateToISO8601(DueDateFrom)
        End If
        If DueDateTo <> Nothing Then
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

    <ExcelFunction(Description:="Function for get receivable", IsThreadSafe:=True)>
    Function AbraReceivable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = Nothing, Optional DocDateTo As Date = Nothing, Optional DueDateFrom As Date = Nothing, Optional DueDateTo As Date = Nothing, Optional Firms As String = "", Optional ACurrency As String = "") As Double
        AbraReceivable = GetReceivable(Url, Username, Password, InformationType, DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency)
    End Function

    <ExcelFunction(Description:="Function for get payable", IsThreadSafe:=True)>
    Function GetPayable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = Nothing, Optional DocDateTo As Date = Nothing, Optional DueDateFrom As Date = Nothing, Optional DueDateTo As Date = Nothing, Optional Firms As String = "", Optional ACurrency As String = "") As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/payable"
        mURL = mURL & "?information-type=" & InformationType
        If DocDateFrom <> Nothing Then
            mURL = mURL & "&doc-date-from=" & DateToISO8601(DocDateFrom)
        End If
        If DocDateTo <> Nothing Then
            mURL = mURL & "&doc-date-to=" & DateToISO8601(DocDateTo)
        End If
        If DueDateFrom <> Nothing Then
            mURL = mURL & "&due-date-from=" & DateToISO8601(DueDateFrom)
        End If
        If DueDateTo <> Nothing Then
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

    <ExcelFunction(Description:="Function for get payable", IsThreadSafe:=True)>
    Function AbraPayable(Url As String, Username As String, Password As String, InformationType As String, Optional DocDateFrom As Date = Nothing, Optional DocDateTo As Date = Nothing, Optional DueDateFrom As Date = Nothing, Optional DueDateTo As Date = Nothing, Optional Firms As String = "", Optional ACurrency As String = "") As Double
        AbraPayable = GetPayable(Url, Username, Password, InformationType, DocDateFrom, DocDateTo, DueDateFrom, DueDateTo, Firms, ACurrency)
    End Function

    <ExcelFunction(Description:="Function for get stock", IsThreadSafe:=True)>
    Function GetStock(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "") As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/stock"
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

    <ExcelFunction(Description:="Function for get stock", IsThreadSafe:=True)>
    Function AbraStock(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "") As Double
        AbraStock = GetStock(Url, Username, Password, InformationType, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores)
    End Function

    <ExcelFunction(Description:="Function for get moves", IsThreadSafe:=True)>
    Function GetMoves(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Object = False, Optional BusOrders As String = "", Optional BusOrdersWithChildren As Object = False, Optional BusTransactions As String = "", Optional BusTransactionsWithChildren As Object = False, Optional BusProjects As String = "", Optional BusProjectsWithChildren As Object = False, Optional Firms As String = "") As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/moves"
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
            If convertBoolean(DivisionsWithChildren) Then
                mURL = mURL & "&divisions-with-children=true"
            End If
        End If
        If BusOrders <> "" Then
            mURL = mURL & "&busorders=" & BusOrders
            If convertBoolean(BusOrdersWithChildren) Then
                mURL = mURL & "busorders-with-children=true"
            End If
        End If
        If BusTransactions <> "" Then
            mURL = mURL & "&bustransactions=" & BusTransactions
            If convertBoolean(BusTransactionsWithChildren) Then
                mURL = mURL & "bustransactions-with-children=true"
            End If
        End If
        If BusProjects <> "" Then
            mURL = mURL & "&busprojects=" & BusProjects
            If convertBoolean(BusProjectsWithChildren) Then
                mURL = mURL & "busprojects-with-children=true"
            End If
        End If
        Dim mResult As String
        mResult = SendRequest(mURL, Username, Password)
        GetMoves = Val(mResult)
    End Function

    <ExcelFunction(Description:="Function for get moves", IsThreadSafe:=True)>
    Function AbraMoves(Url As String, Username As String, Password As String, InformationType As String, DateFrom As Date, DateTo As Date, Optional StoreMenuItems As String = "", Optional StoreCardCategories As String = "", Optional StoreCards As String = "", Optional Stores As String = "", Optional Divisions As String = "", Optional BusOrders As String = "", Optional BusTransactions As String = "", Optional BusProjects As String = "", Optional Firms As String = "") As Double
        AbraMoves = GetMoves(Url, Username, Password, InformationType, DateFrom, DateTo, StoreMenuItems, StoreCardCategories, StoreCards, Stores, Divisions, False, BusOrders, False, BusTransactions, False, BusProjects, False, Firms)
    End Function

    <ExcelFunction(Description:="Function for get depreciation", IsThreadSafe:=True)>
    Function GetDepreciation(Url As String, Username As String, Password As String, InformationType As String, Optional DateFrom As Date = Nothing, Optional DateTo As Date = Nothing, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional AssetLocationsWithChildren As Object = False, Optional Responsibles As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Object = False) As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/depreciation"
        mURL = mURL & "?information-type=" & InformationType
        If DateFrom <> Nothing Then
            mURL = mURL & "&date-from=" & DateToISO8601(DateFrom)
        End If
        If DateTo <> Nothing Then
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
            If convertBoolean(AssetLocationsWithChildren) Then
                mURL = mURL & "&asset-locations-with-children=true"
            End If
        End If
        If Divisions <> "" Then
            mURL = mURL & "&divisions=" & Divisions
            If convertBoolean(DivisionsWithChildren) Then
                mURL = mURL & "&divisions-with-children=true"
            End If
        End If
        Dim mResult As String
        mResult = SendRequest(mURL, Username, Password)
        GetDepreciation = Val(mResult)
    End Function

    <ExcelFunction(Description:="Function for get depreciation", IsThreadSafe:=True)>
    Function AbraDepreciation(Url As String, Username As String, Password As String, InformationType As String, Optional DateFrom As Date = Nothing, Optional DateTo As Date = Nothing, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional Responsibles As String = "", Optional Divisions As String = "") As Double
        AbraDepreciation = GetDepreciation(Url, Username, Password, InformationType, DateFrom, DateTo, AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, False, Responsibles, Divisions, False)
    End Function

    <ExcelFunction(Description:="Function for get asset", IsThreadSafe:=True)>
    Function GetAsset(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional AssetLocationsWithChildren As Object = False, Optional Responsibles As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Object = False) As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/asset"
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
            If convertBoolean(AssetLocationsWithChildren) Then
                mURL = mURL & "&asset-locations-with-children=true"
            End If
        End If
        If Divisions <> "" Then
            mURL = mURL & "&divisions=" & Divisions
            If convertBoolean(DivisionsWithChildren) Then
                mURL = mURL & "&divisions-with-children=true"
            End If
        End If
        Dim mResult As String
        mResult = SendRequest(mURL, Username, Password)
        GetAsset = Val(mResult)
    End Function

    <ExcelFunction(Description:="Function for get asset", IsThreadSafe:=True)>
    Function AbraAsset(Url As String, Username As String, Password As String, InformationType As String, DateTo As Date, Optional AssetTypes As String = "", Optional AccDepreciationGroups As String = "", Optional TaxDepreciationGroups As String = "", Optional AssetLocations As String = "", Optional Responsibles As String = "", Optional Divisions As String = "") As Double
        AbraAsset = GetAsset(Url, Username, Password, InformationType, DateTo, AssetTypes, AccDepreciationGroups, TaxDepreciationGroups, AssetLocations, False, Responsibles, Divisions, False)
    End Function

    Function GetPayroll(Url As String, Username As String, Password As String, InformationType As String, WagePeriods As Object, Optional EmployPatterns As String = "", Optional Divisions As String = "", Optional DivisionsWithChildren As Object = False) As Double
        Dim mURL As String
        mURL = FilterBaseUrl(Url) & "/utils/payroll"
        mURL = mURL & "?information-type=" & InformationType
        Dim mDate As Date = Date.FromOADate(DirectCast(WagePeriods, Double))
        mURL = mURL & "&wage-periods=" & mDate.ToString("yyyy\/MM")
        If EmployPatterns <> "" Then
            mURL = mURL & "&employ-patterns=" & EmployPatterns
        End If
        If Divisions <> "" Then
            mURL = mURL & "&divisions=" & Divisions
            If convertBoolean(DivisionsWithChildren) Then
                mURL = mURL & "&divisions-with-children=true"
            End If
        End If
        Dim mResult As String
        mResult = SendRequest(mURL, Username, Password)
        GetPayroll = Val(mResult)
    End Function

    <ExcelFunction(Description:="Function for get turnover", IsThreadSafe:=True)>
    Function AbraPayroll(Url As String, Username As String, Password As String, InformationType As String, WagePeriods As Object, Optional EmployPatterns As String = "", Optional Divisions As String = "") As Double
        AbraPayroll = GetPayroll(Url, Username, Password, InformationType, WagePeriods, EmployPatterns, Divisions, False)
    End Function
End Module