VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' IWebAuthenticator v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Interface for creating authenticators for rest client
'
' @class IWebAuthenticator
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit


Public Sub BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
End Sub

Public Sub AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
End Sub

Public Sub PrepareHttp(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Http As Object)
End Sub

Public Sub PrepareCurl(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
End Sub
