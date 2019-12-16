VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAddress"
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

Public ADDRESS_ID As Long
Public HOME As String
Public SOI As String
Public MOO As String
Public VILLAGE As String
Public ROAD As String
Public DISTRICT As String
Public AMPHUR As String
Public PROVINCE As String
Public COUNTRY_ID As Long
Public PHONE1 As String
Public PHONE2 As String
Public FAX1 As String
Public FAX2 As String
Public ZIPCODE As String
Public BANGKOK_FLAG As String

Public COUNTRY_NAME As String
Public ADDRESS_TYPE As Long
Public ENTERPRISE_ID As Long
Public CUSTOMER_ID As Long
Public SUPPLIER_ID As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   BANGKOK_FLAG = "N"
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

   SelectStr1 = "SELECT ADR.*, CT.COUNTRY_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ADDRESS ADR, COUNTRY CT "
   
   WhereStr = "WHERE (ADR.COUNTRY_ID = CT.COUNTRY_ID) "
   If ADDRESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADDRESS_ID = " & ADDRESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADDRESS_ID = " & ADDRESS_ID & ") "
      End If
   End If
   
'   If COUNTRY_ID >= 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (COUNTRY_ID = " & COUNTRY_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (COUNTRY_ID = " & COUNTRY_ID & ") "
'      End If
'   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY ADDRESS_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY ADDRESS_ID "
      Else
         OrderByStr = "ORDER BY ADDRESS_ID "
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

   SelectStr1 = "SELECT ADR.*, CT.COUNTRY_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ADDRESS ADR, COUNTRY CT, ENTERPRISE_ADDRESS EA "
   
   WhereStr = "WHERE (ADR.COUNTRY_ID = CT.COUNTRY_ID) AND (ADR.ADDRESS_ID = EA.ADDRESS_ID) "
   If ADDRESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADDRESS_ID = " & ADDRESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADDRESS_ID = " & ADDRESS_ID & ") "
      End If
   End If
   
   If ENTERPRISE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      End If
   End If
   
'   If COUNTRY_ID >= 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (COUNTRY_ID = " & COUNTRY_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (COUNTRY_ID = " & COUNTRY_ID & ") "
'      End If
'   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY ADDRESS_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY ADDRESS_ID "
      Else
         OrderByStr = "ORDER BY ADDRESS_ID "
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

Public Function QueryData3(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT ADR.*, CT.COUNTRY_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ADDRESS ADR, COUNTRY CT, CUSTOMER_ADDRESS CA "
   
   WhereStr = "WHERE (ADR.COUNTRY_ID = CT.COUNTRY_ID) AND (ADR.ADDRESS_ID = CA.ADDRESS_ID) "
   If ADDRESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADDRESS_ID = " & ADDRESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADDRESS_ID = " & ADDRESS_ID & ") "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
'   If COUNTRY_ID >= 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (COUNTRY_ID = " & COUNTRY_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (COUNTRY_ID = " & COUNTRY_ID & ") "
'      End If
'   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ADDRESS_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ADDRESS_ID "
   Else
      OrderByStr = "ORDER BY ADDRESS_ID "
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

Public Function QueryData4(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT ADR.*, CT.COUNTRY_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ADDRESS ADR, COUNTRY CT, SUPPLIER_ADDRESS CA "
   
   WhereStr = "WHERE (ADR.COUNTRY_ID = CT.COUNTRY_ID) AND (ADR.ADDRESS_ID = CA.ADDRESS_ID) "
   If ADDRESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADDRESS_ID = " & ADDRESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADDRESS_ID = " & ADDRESS_ID & ") "
      End If
   End If
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
   
'   If COUNTRY_ID >= 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (COUNTRY_ID = " & COUNTRY_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (COUNTRY_ID = " & COUNTRY_ID & ") "
'      End If
'   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ADDRESS_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ADDRESS_ID "
   Else
      OrderByStr = "ORDER BY ADDRESS_ID "
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

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
   HOME = NVLS(Rs("HOME"), "")
   SOI = NVLS(Rs("SOI"), "")
   MOO = NVLS(Rs("MOO"), "")
   VILLAGE = NVLS(Rs("VILLAGE"), "")
   ROAD = NVLS(Rs("ROAD"), "")
   DISTRICT = NVLS(Rs("DISTRICT"), "")
   AMPHUR = NVLS(Rs("AMPHUR"), "")
   PROVINCE = NVLS(Rs("PROVINCE"), "")
   COUNTRY_ID = NVLI(Rs("COUNTRY_ID"), -1)
   PHONE1 = NVLS(Rs("PHONE1"), "")
   PHONE2 = NVLS(Rs("PHONE2"), "")
   FAX1 = NVLS(Rs("FAX1"), "")
   FAX2 = NVLS(Rs("FAX2"), "")
   ZIPCODE = NVLS(Rs("ZIPCODE"), "")
   COUNTRY_NAME = NVLS(Rs("COUNTRY_NAME"), "")
   BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "N")
End Function

Public Function PackAddress() As String
Dim AddressStr As String

   AddressStr = ""
   
   If Trim(HOME) <> "" Then
      AddressStr = AddressStr & HOME & " "
   End If

   If Trim(SOI) <> "" Then
      AddressStr = AddressStr & "´." & SOI & " "
   End If

   If Trim(MOO) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË") & MOO & " "
   End If

   If Trim(VILLAGE) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË∫È“π") & VILLAGE & " "
   End If

   If Trim(ROAD) <> "" Then
      AddressStr = AddressStr & MapText("∂.") & ROAD & " "
   End If

   If Trim(DISTRICT) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("·¢«ß") & DISTRICT & " "
      Else
         AddressStr = AddressStr & MapText("µ”∫≈") & DISTRICT & " "
      End If
   End If
   
   If Trim(AMPHUR) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("‡¢µ") & AMPHUR & " "
      Else
         AddressStr = AddressStr & MapText("Õ”‡¿Õ") & AMPHUR & " "
      End If
   End If

   If Trim(PROVINCE) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & "" & PROVINCE & " "
      Else
         AddressStr = AddressStr & MapText("®.") & PROVINCE & " "
      End If
   End If

   If Trim(ZIPCODE) <> "" Then
      AddressStr = AddressStr & " " & ZIPCODE & " "
   End If

   PackAddress = AddressStr
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As Long
Dim TempHomeNo As String

   TempID1 = COUNTRY_ID
   If COUNTRY_ID <= 0 Then
      Dim TempCol As Collection
      Set TempCol = New Collection
'      Call LoadCountry(Nothing, TempCol)
      COUNTRY_ID = TempCol(1).COUNTRY_ID
      TempID1 = COUNTRY_ID
      Set TempCol = Nothing
   End If
   
   TempHomeNo = HOME
   If Len(Trim(HOME)) <= 0 Then
      TempHomeNo = "Dummy Home NO."
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("ADDRESS_SEQ", ID, glbErrorLog)
   
      SQL1 = "INSERT INTO ADDRESS " & _
                     "( " & _
                     "   ADDRESS_ID, " & _
                     "   HOME, " & _
                     "   SOI, " & _
                     "   MOO, " & _
                     "   VILLAGE, " & _
                     "   ROAD, " & _
                     "   DISTRICT, " & _
                     "   AMPHUR, " & _
                     "   PROVINCE, " & _
                     "   COUNTRY_ID, " & _
                     "   PHONE1, " & _
                     "   PHONE2, " & _
                     "   FAX1, " & _
                     "   FAX2, " & _
                     "   ZIPCODE, " & _
                     "   BANGKOK_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(TempHomeNo)) & "'," & _
                         "'" & ChangeQuote(Trim(SOI)) & "'," & _
                         "'" & ChangeQuote(Trim(MOO)) & "'," & _
                         "'" & ChangeQuote(Trim(VILLAGE)) & "'," & _
                         "'" & ChangeQuote(Trim(ROAD)) & "'," & _
                         "'" & ChangeQuote(Trim(DISTRICT)) & "'," & _
                         "'" & ChangeQuote(Trim(AMPHUR)) & "'," & _
                         "'" & ChangeQuote(Trim(PROVINCE)) & "'," & _
                         "" & TempID1 & "," & _
                         "'" & ChangeQuote(Trim(PHONE1)) & "'," & _
                         "'" & ChangeQuote(Trim(PHONE2)) & "'," & _
                         "'" & ChangeQuote(Trim(FAX1)) & "'," & _
                         "'" & ChangeQuote(Trim(FAX2)) & "'," & _
                         "'" & ChangeQuote(Trim(ZIPCODE)) & "'," & _
                         "'" & ChangeQuote(Trim(BANGKOK_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         ADDRESS_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ADDRESS SET " & _
                     "   HOME = '" & ChangeQuote(Trim(TempHomeNo)) & "', " & _
                     "   SOI = '" & ChangeQuote(Trim(SOI)) & "', " & _
                     "   MOO = '" & ChangeQuote(Trim(MOO)) & "', " & _
                     "   VILLAGE = '" & ChangeQuote(Trim(VILLAGE)) & "', " & _
                     "   ROAD = '" & ChangeQuote(Trim(ROAD)) & "', " & _
                     "   DISTRICT = '" & ChangeQuote(Trim(DISTRICT)) & "', " & _
                     "   AMPHUR = '" & ChangeQuote(Trim(AMPHUR)) & "', " & _
                     "   PROVINCE = '" & ChangeQuote(Trim(PROVINCE)) & "', " & _
                     "   COUNTRY_ID = " & TempID1 & ", " & _
                     "   PHONE1 = '" & ChangeQuote(Trim(PHONE1)) & "', " & _
                     "   PHONE2 = '" & ChangeQuote(Trim(PHONE2)) & "', " & _
                     "   FAX1 = '" & ChangeQuote(Trim(FAX1)) & "', " & _
                     "   FAX2 = '" & ChangeQuote(Trim(FAX2)) & "', " & _
                     "   ZIPCODE = '" & ChangeQuote(Trim(ZIPCODE)) & "', " & _
                     "   BANGKOK_FLAG = '" & ChangeQuote(Trim(BANGKOK_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ADDRESS_ID = " & ADDRESS_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM ADDRESS WHERE ADDRESS_ID = " & ADDRESS_ID
      m_Conn.Execute (SQL1)
End Function

