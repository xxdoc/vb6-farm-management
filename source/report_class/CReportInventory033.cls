VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory033"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory033"
Public MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Headers3 As CFieldList
Private m_Headers4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ExpenseTypes As Collection
Private m_PartGroup As Collection
Private m_ExpenseTypesEx As Collection
Private m_ExpenseTypes1 As Collection
Private m_ExpenseTypes2 As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_ExpenseTypes = New Collection
   Set m_PartGroup = New Collection
   Set m_ExpenseTypesEx = New Collection
   Set m_ExpenseTypes1 = New Collection
   Set m_ExpenseTypes2 = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Headers3 = New CFieldList
   Set m_Headers4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_ExpenseTypes = Nothing
   Set m_PartGroup = Nothing
   Set m_ExpenseTypesEx = Nothing
   Set m_ExpenseTypes1 = Nothing
   Set m_ExpenseTypes2 = Nothing
      
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Headers3 = Nothing
   Set m_Headers4 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Details2 = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.Add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property

Private Function CReportInterface_Preview() As Boolean
   CReportInterface_Preview = genDoc(True)
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim I As Long
   
   For I = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next I

End Sub
'============================= Codes above used in every report ==========================

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim Sgi As CSGroupItem
Dim TempSgi As CSGroupItem
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Inv As CInventoryDoc
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

   Call LoadExpenseType(Nothing, m_ExpenseTypes1, "Y")
   Call LoadExpenseType(Nothing, m_ExpenseTypes2, "N")
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField
   
   Set Cf = New CReportField
      
   Call Cf.SetFieldValue(1.5, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "จำนวนสุกร (ตัว)", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue((3 * m_PartGroup.Count + 3 * m_ExpenseTypes1.Count + 3 * m_ExpenseTypes2.Count), "^", "มูลค่าต้นทุนเคลื่อนไหว", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ต้นทุน", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ต.ท./ตัว", "^")
   Call m_Headers0.AddField(Cf)
            
   Call Cf.SetFieldValue(1.5, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คลอด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับเข้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จ่ายออก", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "รร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "ส.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", "^")
   Call m_Headers1.AddField(Cf)
   
   Dim Et As CExpenseType
   For Each Et In m_ExpenseTypes1
      Call Cf.SetFieldValue(3, "^", Et.EXPENSE_TYPE_NAME, ">", "E-" & Et.EXPENSE_TYPE_ID)
      Call m_Headers1.AddField(Cf)
   Next Et
   
   Dim Pg As CPartGroup
   For Each Pg In m_PartGroup
      Call Cf.SetFieldValue(3, "^", Pg.PART_GROUP_NAME, ">", "I-" & Pg.PART_GROUP_ID)
      Call m_Headers1.AddField(Cf)
   Next Pg
      
   For Each Et In m_ExpenseTypes2
      Call Cf.SetFieldValue(3, "^", Et.EXPENSE_TYPE_NAME, ">", "E-" & Et.EXPENSE_TYPE_ID)
      Call m_Headers1.AddField(Cf)
   Next Et
   Call Cf.SetFieldValue(3, "^", "ต้นทุน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ต.ท./ตัว", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
      
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)

  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
   
   Call LoadPartGroup(Nothing, m_PartGroup)
   Call LoadExpenseType(Nothing, m_ExpenseTypesEx)
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprFanfoldUS 'pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "จาก วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "  ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim O As Object
Dim RName As String
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Locations As Collection
Dim Pigs As Collection
Dim Lc As CLocation
Dim Pg As CPartItem
Dim InventoryBals As Collection
Dim CapitalBals As Collection
Dim Mi As CMovementItem
Dim Cm As CCapitalMovement
Dim TempID As Long
Dim NewDate As Date
Dim LocationMaps As Collection
Dim StatusMaps As Collection
Dim Ps As CProductStatus
Dim Lm As CLocation
Dim LeftAmount As Double
Dim CapitalMovements As Collection
Dim TempMi As CMovementItem
Dim Sum1 As Double
Dim TotalCount As Long
Dim ImportPigs As Collection
Dim ExportPigs As Collection
Dim II As CImportItem
Dim EI As CExportItem
Dim ParentFlag As String

Dim HaveValue As Boolean            'มีค่าของรายการลูก
Dim HaveValueHeader As Boolean            'มีค่าของรายการลูก
Dim HaveValueGroup As Boolean            'แสดงหัวแล้วรึยัง
   RName = "genDoc"
   
   TotalCount = 0
   
   If mcolParam("PARENT_FLAG") = "Y" Then
      ParentFlag = ""
   Else
      ParentFlag = "N"
   End If
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
   
   Set LocationMaps = New Collection
   Call LoadLocation(Nothing, LocationMaps, 1, "")
   
   Set StatusMaps = New Collection
   Call LoadProductStatus(Nothing, StatusMaps)

   If mcolParam("FROM_DATE") > 0 Then
      Set ImportPigs = New Collection
      Call LoadHousePigImportAmount(Nothing, ImportPigs, -1, NewDate, CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"))
      
      Set ExportPigs = New Collection
      Call LoadPigHouseExportAmountEx(Nothing, ExportPigs, -1, NewDate, CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"))
   End If

   Set CapitalBals = New Collection
   Call LoadCapitalBalanceEx(Nothing, CapitalBals, -1, NewDate, , mcolParam("HOUSE_ID"))
   
   Set CapitalMovements = New Collection
   Call LoadCapitalMovementEx(Nothing, CapitalMovements, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   
   Set Locations = New Collection
   Set Pigs = New Collection
   
   Set Rs = New ADODB.Recordset
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField

      Call LoadMovementLocation(Nothing, Locations, , , mcolParam("HOUSE_ID"), mcolParam("YEAR_SEQ_ID"), mcolParam("WEEK_NO"), PigTypeToCode(mcolParam("PIG_TYPE_ID")))
      For Each Lc In Locations
         For j = 1 To UBound(Total1)
            Total2(j) = 0
         Next j

         Call LoadMovementPig(Nothing, Pigs, , , , Lc.LOCATION_ID, ParentFlag, mcolParam("YEAR_SEQ_ID"), mcolParam("WEEK_NO"), PigTypeToCode(mcolParam("PIG_TYPE_ID")))
         For Each Pg In Pigs
            DoEvents
            Sum1 = 0
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j

'            Call LoadCapitalBalance(Nothing, CapitalBals, -1, NewDate, Pg.PART_ITEM_ID, Lc.LOCATION_ID)
'            Call LoadCapitalMovement(Nothing, CapitalMovements, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), Pg.PART_ITEM_ID, Lc.LOCATION_ID)
            
            'ขึ้น week ใหม่
            Call m_Headers2.ClearField
            Call BodyCf.SetFieldValue(7.5, "^", Lc.LOCATION_NAME & " (" & Lc.LOCATION_NO & ")", "^")
            Call m_Headers2.AddField(BodyCf)
            Call BodyCf.SetFieldValue(10, "^", Pg.PART_NO & " (" & Pg.PIG_TYPE & ")", "^")
            Call m_Headers2.AddField(BodyCf)
            Call BodyCf.SetFieldValue((3 * m_PartGroup.Count + 3 * m_ExpenseTypes1.Count + 3 * m_ExpenseTypes2.Count), "^", "", "^")
            Call m_Headers2.AddField(BodyCf)
            Call BodyCf.SetFieldValue(5, "^", "", "^")
            Call m_Headers2.AddField(BodyCf)
            
            HaveValueGroup = False
            'ขึ้น week ใหม่
            
            'ยอดยกมา
            Call m_Headers3.ClearField
            Set II = GetImportItem(ImportPigs, Lc.LOCATION_ID & "-" & Pg.PART_ITEM_ID)
            Set EI = GetExportItem(ExportPigs, Lc.LOCATION_ID & "-" & Pg.PART_ITEM_ID)
            
            LeftAmount = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT
            j = 0
            HaveValueHeader = False
            For Each HeadCf In m_Headers1.Fields
               j = j + 1
               If j = 2 Then
                  TempStr = "*** ยกมา ***"
                  Amt = 0
               ElseIf j = 9 Then
                  Amt = LeftAmount 'O.CURRENT_AMOUNT
                  TempStr = FormatNumber(Amt, "-")
               Else
                  If Mid(HeadCf.TempKey, 1, 2) = "I-" Then
                     TempID = Mid(HeadCf.TempKey, 3) 'Part group id
                     Set Mi = GetMovementItem(CapitalBals, Lc.LOCATION_ID & "-" & Pg.PART_ITEM_ID & "-" & TempID & "-" & 0)
                     Amt = Mi.CAPITAL_AMOUNT
                     TempStr = FormatNumber(Amt, "-")
                     Sum1 = Sum1 + Amt
                  ElseIf Mid(HeadCf.TempKey, 1, 2) = "E-" Then
                     TempID = Mid(HeadCf.TempKey, 3) 'Expense type
                     Set Mi = GetMovementItem(CapitalBals, Lc.LOCATION_ID & "-" & Pg.PART_ITEM_ID & "-" & 0 & "-" & TempID)
                     Amt = Mi.CAPITAL_AMOUNT
                     TempStr = FormatNumber(Amt, "-")
                     Sum1 = Sum1 + Amt
                  ElseIf j = m_Headers1.Fields.Count - 1 Then
                     Amt = Sum1
                     TempStr = FormatNumber(Amt, "-")
                  ElseIf j = m_Headers1.Fields.Count Then
                     Amt = MyDiffEx(Sum1, LeftAmount)
                     TempStr = FormatNumber(Amt, "-")
                  Else
                     Amt = 0
                     TempStr = ""
                  End If
               End If
               
               If (j = 9) Or (j = 20) Or (j = 21) Then
                  Total1(j) = Amt
               Else
                  Total1(j) = Total1(j) + Amt
               End If
               HaveValueHeader = CheckHaveValue(HaveValueHeader, Amt)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
               Call m_Headers3.AddField(BodyCf)
            Next HeadCf
            
            If HaveValueHeader Then
               If Not HaveValueGroup Then
                  Call m_Headers2.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  'ขึ้น week ใหม่
                  HaveValueGroup = True
               End If
               Call m_Headers3.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            'ยอดยกมา

            'รายการเคลื่อนไหว
            Set Cm = New CCapitalMovement
            Cm.PIG_ID = Pg.PART_ITEM_ID
            Cm.FROM_HOUSE_ID = Lc.LOCATION_ID
            Cm.FROM_DATE = mcolParam("FROM_DATE")
            Cm.TO_DATE = mcolParam("TO_DATE")
            Cm.OrderType = 1
            Call Cm.QueryData(4, Rs, iCount)
            While Not Rs.EOF
               I = I + 1
'If i = 21 Then
'''debug.print
'End If
               Call m_Details1.ClearField
               
               Call Cm.PopulateFromRS(4, Rs)
               
               j = 0
               HaveValue = False
               For Each HeadCf In m_Headers1.Fields
                  j = j + 1
                  If j = 1 Then
                     TempStr = I
                     Amt = 0
                  ElseIf j = 2 Then
                     TempStr = DateToStringExtEx2(Cm.DOCUMENT_DATE)
                     Amt = 0
                  ElseIf j = 3 Then
                     TempStr = Cm.DOCUMENT_NO
                     Amt = 0
                  ElseIf j = 4 Then
                     If (Cm.DOCUMENT_CATEGORY = 1) And (Cm.DOCUMENT_TYPE = 5) Then   'คลอด
                        Amt = Cm.TX_AMOUNT
                        TempStr = FormatNumber(Amt, "-")
                        LeftAmount = LeftAmount + Cm.TX_AMOUNT
                     Else
                        Amt = 0
                        TempStr = FormatNumber(Amt, "-")
                     End If
                  ElseIf j = 5 Then
                     If (Cm.DOCUMENT_CATEGORY = 1) And _
                         (Cm.TX_TYPE = "I") And _
                         ((Cm.DOCUMENT_TYPE = 6) Or (Cm.DOCUMENT_TYPE = 7) Or (Cm.DOCUMENT_TYPE = 8) Or (Cm.DOCUMENT_TYPE = 11) Or (Cm.DOCUMENT_TYPE = 12)) Then 'โอนเข้า
                        Amt = Cm.TX_AMOUNT
                        TempStr = FormatNumber(Amt, "-")
                        
                        LeftAmount = LeftAmount + Cm.TX_AMOUNT
                     Else
                        Amt = 0
                        TempStr = FormatNumber(Amt, "-")
                     End If
                  ElseIf j = 6 Then
                     If (Cm.DOCUMENT_CATEGORY = 1) And _
                         (Cm.TX_TYPE = "E") And _
                         ((Cm.DOCUMENT_TYPE = 6) Or (Cm.DOCUMENT_TYPE = 7) Or (Cm.DOCUMENT_TYPE = 8) Or (Cm.DOCUMENT_TYPE = 9) Or (Cm.DOCUMENT_TYPE = 10) Or (Cm.DOCUMENT_TYPE = 11) Or (Cm.DOCUMENT_TYPE = 12) Or (Cm.DOCUMENT_TYPE = 13)) Then 'โอนออก
                        Amt = Cm.TX_AMOUNT
                        TempStr = FormatNumber(Amt, "-")
                        
                        LeftAmount = LeftAmount - Cm.TX_AMOUNT
                     Else
                        Amt = 0
                        TempStr = FormatNumber(Amt, "-")
                     End If
                  ElseIf j = 7 Then
                     If Cm.TO_HOUSE_ID > 0 Then
                        Set Lm = LocationMaps(Trim(Str(Cm.TO_HOUSE_ID)))
                        TempStr = Lm.LOCATION_NO
                     Else
                        TempStr = ""
                     End If
                     Amt = 0
                  ElseIf j = 8 Then
                     If Cm.PIG_STATUS > 0 Then
                        Set Ps = StatusMaps(Trim(Str(Cm.PIG_STATUS)))
                        TempStr = Ps.PRODUCT_STATUS_NO
                     Else
                        TempStr = ""
                     End If
                     Amt = 0
                  ElseIf j = 9 Then
                     TempStr = FormatNumber(LeftAmount, "-")
                     Amt = LeftAmount
                  ElseIf Mid(HeadCf.TempKey, 1, 2) = "I-" Then
                     TempID = Mid(HeadCf.TempKey, 3) 'Part group id
                     Cm.PART_GROUP_ID = TempID
                     Cm.EXPENSE_TYPE = 0
'''debug.print "2:" & MI.GetKey1
                     Set TempMi = GetMovementItem(CapitalMovements, Lc.LOCATION_ID & "-" & Pg.PART_ITEM_ID & "-" & Cm.GetKey1)
'For Each MI In CapitalMovements
'''debug.print MI.GetKey1
'Next MI
                     Amt = TempMi.CAPITAL_AMOUNT
                     TempStr = FormatNumber(Amt, "-")
                     Sum1 = Sum1 + Amt
                  ElseIf Mid(HeadCf.TempKey, 1, 2) = "E-" Then
                     TempID = Mid(HeadCf.TempKey, 3) 'Expense type
                     Cm.EXPENSE_TYPE = TempID
                     Cm.PART_GROUP_ID = 0
                     
                     Set TempMi = GetMovementItem(CapitalMovements, Lc.LOCATION_ID & "-" & Pg.PART_ITEM_ID & "-" & Cm.GetKey1)
                     Amt = TempMi.CAPITAL_AMOUNT
                     TempStr = FormatNumber(Amt, "-")
                     Sum1 = Sum1 + Amt
                  ElseIf j = m_Headers1.Fields.Count - 1 Then
                     Amt = Sum1
                     TempStr = FormatNumber(Amt, "-")
                  ElseIf j = m_Headers1.Fields.Count Then
                     Amt = MyDiffEx(Sum1, LeftAmount)
                     TempStr = FormatNumber(Amt, "-")
                  Else
                     TempStr = ""
                  End If
                  
                  If (j = 9) Or (j = 22) Or (j = 23) Then
                     Total1(j) = Amt
                  Else
                     Total1(j) = Total1(j) + Amt
                  End If
                  
                  HaveValue = CheckHaveValue(HaveValue, Amt)
                  
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               Next HeadCf
               
               If HaveValue Then
                  If Not HaveValueGroup Then
                     Call m_Headers2.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Vsp.FontBold = True
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                     'ขึ้น week ใหม่
                     HaveValueGroup = True
                  End If
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               Else
                  I = I - 1
               End If
               
               Rs.MoveNext
            Wend
            Set Mi = Nothing
            'รายการเคลื่อนไหวทุน
            
            If HaveValueGroup Then
               Call GenerateFooter("รวม", Pg.PART_NO & " (" & Pg.PIG_TYPE & ")", "", Total1, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            
            For j = 1 To UBound(Total2)
               Total2(j) = Total2(j) + Total1(j)
            Next j
         Next Pg
         
         If HaveValueGroup Then
            Total2(m_Headers1.Fields.Count) = MyDiffEx(Total2(m_Headers1.Fields.Count - 1), Total2(9))
            Call GenerateFooter("รวม", Lc.LOCATION_NAME, "", Total2, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
         For j = 1 To UBound(Total3)
            Total3(j) = Total3(j) + Total2(j)
         Next j
      Next Lc

      Total3(m_Headers1.Fields.Count) = MyDiffEx(Total3(m_Headers1.Fields.Count - 1), Total3(9))
      Call GenerateFooter("สุทธิ", "", "", Total3, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   Set CapitalMovements = Nothing
   Set LocationMaps = Nothing
   Set StatusMaps = Nothing
   Set CapitalBals = Nothing
   Set InventoryBals = Nothing
   Set Locations = Nothing
   Set Pigs = Nothing
   Set ImportPigs = Nothing
   Set ExportPigs = Nothing

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf (j = 4) Then
         TempStr = FormatNumber(Tot(j), "-")
      ElseIf (j = 5) Then
         TempStr = FormatNumber(Tot(j), "-")
      ElseIf j = 6 Then
         TempStr = FormatNumber(Tot(j), "-")
      ElseIf j = 7 Then
         TempStr = ""
      ElseIf j = 8 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(j), "-")
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub

Private Sub VSP_NewPage()
Dim talnT As TextAlignSettings
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long

   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|version|: " & glbParameterObj.Version)
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   'Vsp.Paragraph = ""
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
   
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY
   
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub


