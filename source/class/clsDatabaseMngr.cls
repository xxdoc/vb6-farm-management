VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabaseMngr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "clsDatabaseMngr"

Private Enum EVENT_TYPE
   EVENT_NONE = -1
   EVENT_CONNECTED = 0
   EVENT_ERROR = 1
   EVENT_CLOSE = 2
   EVENT_DATARECEIVED = 3
   EVENT_SENTCOMPLETE = 4
End Enum

Private m_LegacyADOConn As ADODB.Connection
Private m_ADOConn As ADODB.Connection
Private m_Rs As ADODB.Recordset
Public WithEvents m_Winsock As Winsock
Attribute m_Winsock.VB_VarHelpID = -1
Private m_Event As EVENT_TYPE
Private m_LastError As String
Private m_StringReceive As String
Private m_HasConnect As Boolean
Private m_IP As String
Private m_Port As Long

Private Type TProvince
   PROVINCE_ID As Long
   PROVINCE_NAME As String
End Type

Private Type TExpenseType
   EXPENSE_TYPE_ID As Long
   EXPENSE_TYPE_NAME As String
End Type

Private m_Province() As TProvince
Private m_ExpenseType() As TExpenseType

Private Sub Class_Initialize()
On Error GoTo ErrorHandler
Dim RName As String

   RName = "Class_Initialize"
   
   Set m_LegacyADOConn = New ADODB.Connection
   Set m_ADOConn = New ADODB.Connection
   Set m_Rs = New ADODB.Recordset
   
   Load frmWinsock
   Set m_Winsock = frmWinsock.Winsock1
   
   Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Property Get DBConnection() As ADODB.Connection
   Set DBConnection = m_ADOConn
End Property

Public Property Get DBLegacyConnection() As ADODB.Connection
   Set DBLegacyConnection = m_LegacyADOConn
End Property

Public Function GetServerDateTime(ServerDateTime As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim ItemCount As Long

   RName = "GetServerDateTime"
   GetServerDateTime = False
   
   SQL = "SELECT CURRENT_TIMESTAMP || '' FROM SYSTEM_PARAM WHERE PARAM_NAME = 'DUMMY_PARAM' "
   If Not GetRs(SQL, "", False, ItemCount, m_Rs, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "End of file condition."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   If NVLS(m_Rs(0), "") = "" Then
      ErrorObj.LocalErrorMsg = "TIMESTAMP is null."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   ServerDateTime = m_Rs(0)
   
   GetServerDateTime = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL -> " & SQL
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetServerDateTime = False
End Function

Public Function GetSeqID(SequenceName As String, ID As Long, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim ItemCount As Long

   RName = "GetSeqID"
   GetSeqID = False
   
   SQL = "SELECT GEN_ID(" & SequenceName & ", 1) FROM SYSTEM_PARAM WHERE PARAM_NAME = 'DUMMY_PARAM' "
   If Not GetRs(SQL, "", False, ItemCount, m_Rs, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "End of file condition."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   If NVLI(m_Rs(0), 0) = 0 Then
      ErrorObj.LocalErrorMsg = "ID is null."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   ID = m_Rs(0)
   
   GetSeqID = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL -> " & SQL
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetSeqID = False
End Function

Public Function GetRs(SQL1 As String, SQL2 As String, MustCount As Boolean, ItemCount As Long, Rs As ADODB.Recordset, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "GetRs"
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   
   End If
   
   If MustCount Then
      Call Rs.Open(SQL2, m_ADOConn, adOpenKeyset, adLockOptimistic) 'มีปัญหาตรงนี้
       If Rs.EOF Then
         ItemCount = 0
      Else
         ItemCount = NVLI(Rs(0), 0)
      End If
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   m_ADOConn.CursorLocation = adUseClient
   Call Rs.Open(SQL1, m_ADOConn, adOpenKeyset, adLockOptimistic) 'มีปัญหาตรงนี้
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   If Rs.State <> adStateOpen Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   GetRs = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetRs = False
End Function

Public Function GetLegacyRs(SQL1 As String, SQL2 As String, MustCount As Boolean, ItemCount As Long, Rs As ADODB.Recordset, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "GetLegacyRs"
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetLegacyRs = False
      Exit Function
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   If MustCount Then
      Call Rs.Open(SQL2, m_LegacyADOConn, adOpenKeyset, adLockOptimistic)
      If Rs.EOF Then
         ItemCount = 0
      Else
         ItemCount = NVLI(Rs(0), 0)
      End If
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   m_LegacyADOConn.CursorLocation = adUseClient
   Call Rs.Open(SQL1, m_LegacyADOConn, adOpenKeyset, adLockOptimistic)
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetLegacyRs = False
      Exit Function
   End If
   
   If Rs.State <> adStateOpen Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetLegacyRs = False
      Exit Function
   End If
   
   GetLegacyRs = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetLegacyRs = False
End Function

Private Function Connet(IP As String, Port As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "Connect"
   Connet = False
#If WIN_DEBUG = 1 Then
   Connet = True
   Exit Function
#End If

   Call m_Winsock.Connect(IP, Port)
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_CONNECTED Then
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.RoutineName = RName
      ErrorObj.LocalErrorMsg = "Can't connect to " & IP & " : " & Port & ""
      ErrorObj.SystemErrorMsg = m_LastError
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      
'      m_Winsock.Close
      Exit Function
   End If
   
   Connet = True
   Exit Function
   
ErrorHandler:
'   m_Winsock.Close
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   Connet = False
End Function

Private Function GetResult(Command As String, FullCommand As String, Result As String, ErrorStr As String) As Boolean
Dim Token As TOKENLIB
Dim Buf As String * 255
Dim tokenstr As String * 255
Dim TokenStrLen As Long
Dim NewToken As String
Dim RName As String

   RName = "GetResult"
   GetResult = False
   
   If TOKInitial(Token) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf
      Exit Function
   End If
   
   TokenStrLen = Len(tokenstr)
   If TOKGetTokenStr(Token, FullCommand, Len(FullCommand), 1, tokenstr, TokenStrLen) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf & "-> " & FullCommand
      Exit Function
   End If
   NewToken = Mid(tokenstr, 1, TokenStrLen)
   If UCase(NewToken) <> "RESPONSE" Then
      Call TOKTerminate(Token)
      ErrorStr = "Not valid valid response command -> " & FullCommand & " "
      Exit Function
   End If
   
   TokenStrLen = Len(tokenstr)
   If TOKGetTokenStr(Token, FullCommand, Len(FullCommand), 2, tokenstr, TokenStrLen) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf & "-> " & "& fullcommand &"
      Exit Function
   End If
   NewToken = Trim(Mid(tokenstr, 1, TokenStrLen))
   If UCase(Command) <> UCase(NewToken) Then
      Call TOKTerminate(Token)
      ErrorStr = "Not valid valid response command -> " & FullCommand & " "
      Exit Function
   End If
   
   TokenStrLen = Len(tokenstr)
   If TOKGetTokenStr(Token, FullCommand, Len(FullCommand), 3, tokenstr, TokenStrLen) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf & "-> " & FullCommand
      Exit Function
   End If
   NewToken = Mid(tokenstr, 1, TokenStrLen)
   If Not IsNumeric(NewToken) Then
      Call TOKTerminate(Token)
      ErrorStr = "Not numeric result value -> " & FullCommand & " "
      Exit Function
   End If
   
   Result = NewToken
   
   TokenStrLen = Len(tokenstr)
   If TOKGetTokenStr(Token, FullCommand, Len(FullCommand), 4, tokenstr, TokenStrLen) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf & "-> " & FullCommand
      Exit Function
   End If
   NewToken = Mid(tokenstr, 1, TokenStrLen)
   
   ErrorStr = NewToken
   
   GetResult = True
End Function

Public Function Login(UserName As String, UserID As String, CanLogin As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim Command As String
Dim Result As String
Dim ErrorStr As String

   RName = "Login"
   Login = False
   
#If WIN_DEBUG = 1 Then
   CanLogin = True
   Login = True
   Exit Function
#End If

   Command = "(login, " & UserName & ", " & UserID & ")"
'   Call m_Winsock.SendData(Command)
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event = EVENT_DATARECEIVED Then
      GoTo XXX
   End If
   If m_Event <> EVENT_SENTCOMPLETE Then
      ErrorObj.LocalErrorMsg = "Send complete event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_DATARECEIVED Then
      ErrorObj.LocalErrorMsg = "Data receive event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
XXX:
   If Not GetResult("login", m_StringReceive, Result, ErrorStr) Then
      ErrorObj.LocalErrorMsg = "Can't extract value from response command."
      ErrorObj.SystemErrorMsg = ErrorStr
      Exit Function
   End If
   
   If Result <> 0 Then
      ErrorObj.LocalErrorMsg = ErrorStr
      ErrorObj.SystemErrorMsg = m_StringReceive
      
      Login = True
      CanLogin = False
      Exit Function
   End If
   
   CanLogin = True
   Login = True
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   
   Login = False
End Function

Private Function IsLogin(UserID As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim Command As String
Dim Result As String
Dim ErrorStr As String

   RName = "IsLogin"
   IsLogin = False
   
   Command = "(islogin,  " & UserID & ")"
'   Call m_Winsock.SendData(Command)
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event = EVENT_DATARECEIVED Then
      GoTo XXX
   End If
   If m_Event <> EVENT_SENTCOMPLETE Then
      ErrorObj.LocalErrorMsg = "Send complete event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_DATARECEIVED Then
      ErrorObj.LocalErrorMsg = "Data receive event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
XXX:
   If Not GetResult("islogin", m_StringReceive, Result, ErrorStr) Then
      ErrorObj.LocalErrorMsg = "Can't extract value from response command."
      ErrorObj.SystemErrorMsg = ErrorStr
      Exit Function
   End If
   
   If Trim(Result) = 0 Then
      IsLogin = False
   Else
      IsLogin = True
   End If
   
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   
   IsLogin = False
End Function

Public Function LockTable(TableName As String, Key As Long, IsCanLock As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim Command As String
Dim Result As String
Dim ErrorStr As String

   RName = "LockTable"
   LockTable = False

#If WIN_DEBUG = 1 Then
   LockTable = True
   IsCanLock = True
   Exit Function
#End If
   
   Command = "(lock, " & TableName & ", " & Key & ")"
'   Call m_Winsock.SendData(Command)
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event = EVENT_DATARECEIVED Then
      GoTo XXX
   End If
   If m_Event <> EVENT_SENTCOMPLETE Then
      ErrorObj.LocalErrorMsg = "Send complete event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_DATARECEIVED Then
      ErrorObj.LocalErrorMsg = "Data receive event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
XXX:
   If Not GetResult("lock", m_StringReceive, Result, ErrorStr) Then
      ErrorObj.LocalErrorMsg = "Can't extract value from response command."
      ErrorObj.SystemErrorMsg = ErrorStr
      Exit Function
   End If
   
   If Result <> 0 Then
      ErrorObj.LocalErrorMsg = ErrorStr
      ErrorObj.SystemErrorMsg = m_StringReceive
      
      LockTable = True
      IsCanLock = False
      Exit Function
   End If
   
   IsCanLock = True
   LockTable = True
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   
   LockTable = False
End Function

Public Function UnLockTable(TableName As String, Key As Long, IsCanUnlock As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim Command As String
Dim Result As String
Dim ErrorStr As String

   RName = "UnLockTable"
   UnLockTable = False
   
#If WIN_DEBUG = 1 Then
   UnLockTable = True
   IsCanUnlock = True
   Exit Function
#End If
   
   Command = "(unlock, " & TableName & ", " & Key & ")"
'   Call m_Winsock.SendData(Command)
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event = EVENT_DATARECEIVED Then
      GoTo XXX
   End If
   If m_Event <> EVENT_SENTCOMPLETE Then
      ErrorObj.LocalErrorMsg = "Send complete event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_DATARECEIVED Then
      ErrorObj.LocalErrorMsg = "Data receive event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
XXX:
   If Not GetResult("unlock", m_StringReceive, Result, ErrorStr) Then
      ErrorObj.LocalErrorMsg = "Can't extract value from response command."
      ErrorObj.SystemErrorMsg = ErrorStr
      Exit Function
   End If
   
   If Result <> 0 Then
      ErrorObj.LocalErrorMsg = ErrorStr
      ErrorObj.SystemErrorMsg = m_StringReceive
      
      UnLockTable = True
      IsCanUnlock = False
      Exit Function
   End If
   
   IsCanUnlock = True
   UnLockTable = True
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   
   UnLockTable = False
End Function

Public Function ConnectDatabase(FileDb As String, UserName As String, Password As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim ConnStr As String

   RName = "ConnectDatabase"

   ConnStr = "data source=" & FileDb & ";ctype=UNICODE_FSS"
'   m_ADOConn.Open ("DSN=TestFireBird;")
   m_ADOConn.Provider = "LCPI.IBProvider"
   Call m_ADOConn.Open(ConnStr, UserName, Password)
   
   m_ADOConn.Properties("Session AutoCommit") = True   ' enable auto commit
   m_ADOConn.Properties("Autocommit Isolation Levels") = 4096  ' read commited=0x1000
      
   ConnectDatabase = True
   Exit Function

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   Call glbErrorLog.ShowErrorLog(LOG_FILE_MSGBOX, False)

   Unload frmSplash
   Set frmSplash = Nothing
   
   ConnectDatabase = False
End Function

Public Function ConnectLegacyDatabase(FileDb As String, UserName As String, Password As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim ConnStr As String

   RName = "ConnectLegacyDatabase"

   ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileDb & ";User Id=;Password=;"
   Call m_LegacyADOConn.Open(ConnStr)
      
   ConnectLegacyDatabase = True
   Exit Function

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   Call glbErrorLog.ShowErrorLog(LOG_FILE_MSGBOX, False)

   Unload frmSplash
   Set frmSplash = Nothing
   
   ConnectLegacyDatabase = False
End Function

Public Function DisConnectLegacyDatabase() As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim ConnStr As String

   RName = "DisConnectLegacyDatabase"
   Call m_LegacyADOConn.Close
      
   DisConnectLegacyDatabase = True
   Exit Function

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   Call glbErrorLog.ShowErrorLog(LOG_FILE_MSGBOX, False)
   
   DisConnectLegacyDatabase = False
End Function

Public Function IsConnect() As Boolean
   IsConnect = m_HasConnect
End Function

Public Function ConnectAgentServer(IP As String, Port As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "ConnectAgentServer"
   ConnectAgentServer = False
   
   If Not Connet(IP, Port, ErrorObj) Then
      Exit Function
   End If
   
   ConnectAgentServer = True
   Exit Function

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

   Unload frmSplash
   Set frmSplash = Nothing
   
   ConnectAgentServer = False
End Function

Private Function ExtractDate(DateStr As String) As String
   ExtractDate = Mid(DateStr, 1, 10)
End Function

Private Function ExtractTime(DateStr As String) As String
   ExtractTime = Mid(DateStr, 12, 8)
End Function

Private Sub Class_Terminate()
   Set m_Winsock = Nothing
   
   Unload frmWinsock
   Set frmWinsock = Nothing
End Sub

Private Sub m_Winsock_Close()
   m_Event = EVENT_CLOSE
   m_HasConnect = False
End Sub

Private Sub m_Winsock_Connect()
   m_Event = EVENT_CONNECTED
   m_HasConnect = True
End Sub

Private Sub m_Winsock_DataArrival(ByVal bytesTotal As Long)
   m_Event = EVENT_DATARECEIVED
'   Call m_Winsock.GetData(m_StringReceive, vbString)
End Sub

Private Sub m_Winsock_Error(ByVal Number As Integer, DESCRIPTION As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   m_Event = EVENT_ERROR
   m_LastError = DESCRIPTION
   m_HasConnect = False
End Sub

Private Sub m_Winsock_SendComplete()
   m_Event = EVENT_SENTCOMPLETE
End Sub

Public Function CountRecord(TableName As String, FieldName1 As String, FieldName2 As String, FieldValue As String, ID As Long, ErrorObj As clsErrorLog, Optional FieldNameExTend As String, Optional FieldNameExTendValue As String) As Long
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim ItemCount As Long

   RName = "CountRecord"
   CountRecord = False
    
    If Len(FieldNameExTend) > 0 Then
        SQL1 = "SELECT COUNT(*) RC FROM " & TableName & " WHERE (" & FieldName1 & " = '" & FieldValue & "') AND (" & FieldName2 & " <> " & ID & ") AND (" & FieldNameExTend & " = '" & FieldNameExTendValue & "')"
   Else
        SQL1 = "SELECT COUNT(*) RC FROM " & TableName & " WHERE (" & FieldName1 & " = '" & FieldValue & "') AND (" & FieldName2 & " <> " & ID & ") "
   End If
   If Not GetRs(SQL1, "", False, ItemCount, m_Rs, ErrorObj) Then
      Exit Function
   End If
   
   If m_Rs.EOF Then
      CountRecord = 0
   Else
      CountRecord = NVLI(m_Rs("RC"), 0)
   End If

   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   CountRecord = False
End Function

'======================== END OF GENERIC FUNCTION HERE ======================

Public Function GenerateNumber(NumberType As NUMBER_TYPE, No As String, ErrorLog As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ServerDateTime As String
Dim ID As Long

   RName = "GenerateNumber"
   GenerateNumber = False
   HasBegin = False

   If NumberType = OPERATE_NUMBER Then
      If Not GetSeqID("OPERATE_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "00000000")
      No = "OP-" & No
   ElseIf NumberType = DO_NUMBER Then
      If Not GetSeqID("DO_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "DO-" & No
   ElseIf NumberType = JOBPLAN_NUMBER Then
      If Not GetSeqID("JOB_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "JP-" & No
   ElseIf NumberType = BORROW_NUMBER Then
      If Not GetSeqID("BORROW_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "BR-" & No
   ElseIf NumberType = DEBIT_NOTE_NUMBER Then
      If Not GetSeqID("BILL_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "BL-" & No
   ElseIf NumberType = QUOATATION_NUMBER Then
      If Not GetSeqID("QUOATATION_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "QT-" & No
   ElseIf NumberType = BILL_NUMBER Then
      If Not GetSeqID("BILLS_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "BS-" & No
   'bum+ start
   ElseIf NumberType = EXPENSE_NUMBER Then
      If Not GetSeqID("EXPENSE_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "CF-" & No
   ElseIf NumberType = REVENUE_NUMBER Then
      If Not GetSeqID("REVENUE_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "RV-" & No
   'bum+ end
   ElseIf NumberType = IMPORT_NUMBER Then
      If Not GetSeqID("IMPORT_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "IMP-" & No
   ElseIf NumberType = EXPORT_NUMBER Then
      If Not GetSeqID("EXPORT_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "EXP-" & No
   ElseIf NumberType = PLAN_NUMBER Then
         If Not GetSeqID("PLAN_KEY_SEQ", ID, ErrorLog) Then
            GenerateNumber = False
            Exit Function
         End If
         No = Format(ID, "0000000")
         No = "PLN-" & No
   ElseIf NumberType = FUEL_NUMBER1 Then
      If Not GetSeqID("FUEL1_KEY_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "OIL-I-" & No
   ElseIf NumberType = FUEL_NUMBER2 Then
      If Not GetSeqID("FUEL2_KEY_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "OIL-O-" & No
   ElseIf NumberType = REPAIR_NUMBER Then
      If Not GetSeqID("REPAIR_KEY_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "REP-" & No
   ElseIf NumberType = QUOATATION_NUMBER Then
      If Not GetSeqID("QUOATATION_KEY_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "QOT-" & No
   ElseIf NumberType = PO_NUMBER Then
      If Not GetSeqID("PO_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "JB-" & No
   ElseIf NumberType = RECEIPT_NUMBER Then
      If Not GetSeqID("RECEIPT_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "RC-" & No
   ElseIf NumberType = INVOICE_RECEIPT_NUMBER Then
      If Not GetSeqID("INVOICE_RECEIPT_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "IR-" & No
   ElseIf NumberType = DBN_NUMBER Then
      If Not GetSeqID("DN_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "DN-" & No
   ElseIf NumberType = CDN_NUMBER Then
      If Not GetSeqID("CN_NUMBER_SEQ", ID, ErrorLog) Then
         GenerateNumber = False
         Exit Function
      End If
      
      No = Format(ID, "0000000")
      No = "CN-" & No
   Else
      No = "NOT DEFINE"
   End If

   GenerateNumber = True
   Exit Function

ErrorHandler:
   ErrorLog.LocalErrorMsg = "Runtime error."
   ErrorLog.SystemErrorMsg = Err.DESCRIPTION
   ErrorLog.RoutineName = RName
   ErrorLog.ModuleName = MODULE_NAME
   ErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

   GenerateNumber = False
End Function
