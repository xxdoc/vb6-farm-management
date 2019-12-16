VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Public SUPPLIER_ID As Long
Public SUPPLIER_CODE As String
Public SUPPLIER_GRADE As Long
Public CREDIT As Long
Public TAX_ID As String
Public SUPPLIER_TYPE As Long
Public EMAIL As String
Public BIRTH_DATE As Date
Public WEBSITE As String
Public PASSWORD1 As String
Public SUPPLIER_STATUS As Long
Public BUSINESS_DESC As String

Public KEY_ID As Long
Public KEY_LOOKUP As String
Public SUPPLIER_NAME As String
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE As Long
Public SUPPLIER_TYPE_NAME As String
Public SUPPLIER_GRADE_NAME As String

Public CstNames As Collection
Public CstDocs As Collection
Public CstAddr As Collection
Public CstContacts As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set CstNames = New Collection
   Set CstDocs = New Collection
   Set CstAddr = New Collection
   Set CstContacts = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT ADR.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SUPPLIER ADR "
   
   If SUPPLIER_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SUPPLIER_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SUPPLIER_ID "
      Else
         OrderByStr = "ORDER BY SUPPLIER_ID "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'In this class we will not handle any error or any transaction
Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT BT.SUPPLIER_STATUS_NAME, CM.*, NM.LONG_NAME SUPPLIER_NAME, CT.SUPPLIER_TYPE_ID, CT.SUPPLIER_TYPE_NAME, CG.SUPPLIER_GRADE_ID, CG.SUPPLIER_GRADE_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = "FROM SUPPLIER CM "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME CN ON (CN.SUPPLIER_ID = CM.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_TYPE CT ON (CM.SUPPLIER_TYPE = CT.SUPPLIER_TYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_GRADE CG ON (CM.SUPPLIER_GRADE = CG.SUPPLIER_GRADE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_STATUS BT ON (CM.SUPPLIER_STATUS = BT.SUPPLIER_STATUS_ID) "
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
      
   If SUPPLIER_GRADE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_GRADE = " & SUPPLIER_GRADE & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_GRADE = " & SUPPLIER_GRADE & ") "
      End If
   End If

   If SUPPLIER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_TYPE = " & SUPPLIER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_TYPE = " & SUPPLIER_TYPE & ") "
      End If
   End If

   If SUPPLIER_STATUS > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_STATUS = " & SUPPLIER_STATUS & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_STATUS = " & SUPPLIER_STATUS & ") "
      End If
   End If

   If Trim(SUPPLIER_NAME) <> "" Then
      If WildCard(SUPPLIER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(SUPPLIER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(SUPPLIER_NAME)) & "')"
         End If
      End If
   End If

   If Trim(SUPPLIER_CODE) <> "" Then
      If WildCard(SUPPLIER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         End If
      End If
   End If

   Dim TempStr As String
   If DOCUMENT_TYPE > 0 Then
      TempStr = "(SELECT COUNT(*) FROM SUPPLIER_PROOF CP, DOCUMENT DM WHERE (CP.DOCUMENT_ID = DM.DOCUMENT_ID) AND (CP.SUPPLIER_ID = CM.SUPPLIER_ID) AND (DM.DOCUMENT_NO = '" & DOCUMENT_NO & "') AND (DM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ")) > 0 "
      If WhereStr = "" Then
         WhereStr = "WHERE (" & TempStr & ") "
      Else
         WhereStr = WhereStr & "AND (" & TempStr & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SUPPLIER_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   Else
      OrderByStr = "ORDER BY CM.SUPPLIER_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
   SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
   SUPPLIER_GRADE = NVLI(Rs("SUPPLIER_GRADE"), -1)
   CREDIT = NVLS(Rs("CREDIT"), "")
   TAX_ID = NVLS(Rs("TAX_ID"), "")
   SUPPLIER_TYPE = NVLI(Rs("SUPPLIER_TYPE"), -1)
   EMAIL = NVLS(Rs("EMAIL"), "")
   WEBSITE = NVLS(Rs("WEBSITE"), "")
   BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
   PASSWORD1 = DecryptText(NVLS(Rs("PASSWORD1"), ""))
   SUPPLIER_STATUS = NVLI(Rs("SUPPLIER_STATUS"), -1)
   BUSINESS_DESC = NVLS(Rs("BUSINESS_DESC"), "")
   
   KEY_LOOKUP = SUPPLIER_CODE
   KEY_ID = SUPPLIER_ID
   SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
   SUPPLIER_GRADE_NAME = NVLS(Rs("SUPPLIER_GRADE_NAME"), "")
   SUPPLIER_TYPE_NAME = NVLS(Rs("SUPPLIER_TYPE_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = SUPPLIER_STATUS
   If SUPPLIER_STATUS <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = SUPPLIER_TYPE
   If SUPPLIER_TYPE <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = SUPPLIER_GRADE
   If SUPPLIER_GRADE <= 0 Then
      TempID3 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SUPPLIER_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO SUPPLIER " & _
                     "( " & _
                     "   SUPPLIER_ID, " & _
                     "   SUPPLIER_CODE, " & _
                     "   SUPPLIER_GRADE, " & _
                     "   CREDIT, " & _
                     "   TAX_ID, " & _
                     "   SUPPLIER_TYPE, " & _
                     "   EMAIL, " & _
                     "   WEBSITE, " & _
                     "   PASSWORD1, " & _
                     "   BIRTH_DATE, " & _
                     "   SUPPLIER_STATUS, " & _
                     "   BUSINESS_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(SUPPLIER_CODE)) & "'," & _
                         TempID3 & "," & _
                         "'" & ChangeQuote(Trim(CREDIT)) & "'," & _
                         "'" & ChangeQuote(Trim(TAX_ID)) & "'," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(EMAIL)) & "'," & _
                         "'" & ChangeQuote(Trim(WEBSITE)) & "'," & _
                         "'" & ChangeQuote(Trim(EncryptText(PASSWORD1))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "'," & _
                         TempID1 & "," & _
                         "'" & ChangeQuote(Trim(BUSINESS_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SUPPLIER_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SUPPLIER SET " & _
                     "   SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "', " & _
                     "   SUPPLIER_GRADE = " & TempID3 & ", " & _
                     "   CREDIT = '" & CREDIT & "', " & _
                     "   TAX_ID = '" & ChangeQuote(Trim(TAX_ID)) & "', " & _
                     "   SUPPLIER_TYPE = " & TempID2 & ", " & _
                     "   EMAIL = '" & ChangeQuote(Trim(EMAIL)) & "', " & _
                     "   WEBSITE = '" & ChangeQuote(Trim(WEBSITE)) & "', " & _
                     "   PASSWORD1 = '" & ChangeQuote(Trim(EncryptText(PASSWORD1))) & "', " & _
                     "   BIRTH_DATE = '" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "', " & _
                     "   SUPPLIER_STATUS = " & TempID1 & ", " & _
                     "   BUSINESS_DESC = '" & ChangeQuote(Trim(BUSINESS_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUPPLIER_ID = " & SUPPLIER_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM SUPPLIER_ADDRESS WHERE SUPPLIER_ID = " & SUPPLIER_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM SUPPLIER_NAME WHERE SUPPLIER_ID = " & SUPPLIER_ID
      m_Conn.Execute (SQL1)
      
'      SQL1 = "DELETE FROM SUPPLIER_PROOF WHERE SUPPLIER_ID = " & SUPPLIER_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM SUPPLIER_CONTACT WHERE SUPPLIER_ID = " & SUPPLIER_ID
'      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM SUPPLIER WHERE SUPPLIER_ID = " & SUPPLIER_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CstNames = Nothing
   Set CstDocs = Nothing
   Set CstAddr = Nothing
   Set CstContacts = Nothing
End Sub
