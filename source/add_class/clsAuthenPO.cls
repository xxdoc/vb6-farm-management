VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAuthenPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsAuthenPO"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing

   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub
'==
Public Function QueryAuthenPO(Ua As CAuthenPO, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryAuthenPO"
   QueryAuthenPO = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
    'CollAuthenPOCheck ++++++++++++++++++++++++++
      Dim Apo As CAuthenPOItem
      Set Apo = New CAuthenPOItem
      Apo.AUTHEN_PO_ID = Ua.AUTHEN_PO_ID
      Apo.AUTHEN_AREA = 1
      Call Apo.QueryData(1, m_Rs2, iCount)
      Set Apo = Nothing

      Set Ua.CollAuthenPOCheck = Nothing
      Set Ua.CollAuthenPOCheck = New Collection
      While Not m_Rs2.EOF
         Set Apo = New CAuthenPOItem
         Call Apo.PopulateFromRS(1, m_Rs2)
         Apo.Flag = "I"
         Call Ua.CollAuthenPOCheck.add(Apo)
         Set Apo = Nothing
         m_Rs2.MoveNext
      Wend
      Set Apo = Nothing
    'CollAuthenPOCheck ++++++++++++++++++++++++++

    'CollAuthenPOApprove ++++++++++++++++++++++++++
      Set Apo = New CAuthenPOItem
      Apo.AUTHEN_PO_ID = Ua.AUTHEN_PO_ID
      Apo.AUTHEN_AREA = 2
      Call Apo.QueryData(1, m_Rs2, iCount)
      Set Apo = Nothing

      Set Ua.CollAuthenPOApprove = Nothing
      Set Ua.CollAuthenPOApprove = New Collection
      While Not m_Rs2.EOF
         Set Apo = New CAuthenPOItem
         Call Apo.PopulateFromRS(1, m_Rs2)
         Apo.Flag = "I"
         Call Ua.CollAuthenPOApprove.add(Apo)
         Set Apo = Nothing
         m_Rs2.MoveNext
      Wend
      Set Apo = Nothing
   'CollAuthenPOApprove ++++++++++++++++++++++++++

   End If

   QueryAuthenPO = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryAuthenPO = False
End Function
Public Function AddEditAuthenPO(Ug As CAuthenPO, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditAuthenPO"
   AddEditAuthenPO = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim Apo As CAuthenPOItem
   For Each Apo In Ug.CollAuthenPOCheck
      If Apo.Flag = "A" Then
         Apo.AUTHEN_PO_ID = Ug.AUTHEN_PO_ID
         Apo.AUTHEN_AREA = 1
         Apo.AddEditMode = SHOW_ADD
         Apo.AddEditData
      ElseIf Apo.Flag = "E" Then
         Apo.AddEditMode = SHOW_EDIT
         Apo.AddEditData
      ElseIf Apo.Flag = "D" Then
         Call Apo.DeleteData
      End If
   Next Apo

  For Each Apo In Ug.CollAuthenPOApprove
      If Apo.Flag = "A" Then
         Apo.AUTHEN_PO_ID = Ug.AUTHEN_PO_ID
         Apo.AUTHEN_AREA = 2
         Apo.AddEditMode = SHOW_ADD
         Apo.AddEditData
      ElseIf Apo.Flag = "E" Then
         Apo.AddEditMode = SHOW_EDIT
         Apo.AddEditData
      ElseIf Apo.Flag = "D" Then
         Call Apo.DeleteData
      End If
   Next Apo

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditAuthenPO = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditAuthenPO = False
End Function

Public Function DeleteAuthenPO(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CAuthenPO

   RName = "DeleteAuthenPO"
   DeleteAuthenPO = False
   HasBegin = False

   Set B = New CAuthenPO

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.AUTHEN_PO_ID = UID

   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteAuthenPO = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteAuthenPO = True
End Function
