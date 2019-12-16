VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory028"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory028"
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

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT1 = "^1|^3|^2|^2|^2|%P^2|^2;"
Private Const HEADER_FORMAT2 = "^1|^3|^2|^2|^2|%P^2|^2;"
Private Const DETAIL_FORMAT1 = "<4|<4|%N;"
Private Const DETAIL_FORMAT2 = "^1|^3|>2|>2|>2|%P>2|>2;"
Private Const FOOTER_FORMAT1 = "^1|^3|>2|>2|%P>2|>2;"
Private Const FOOTER_FORMAT2 = "^1|^3|>2|>2|>2|%P>2|>2;"
'NO.|วันที่|เลขที่เอกสาร|เกิด|โอนเข้า|%P|โอนออก|คงเหลือ;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_YearSeqs = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
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

   Set ImportItems = New Collection
   Set TempRs = New ADODB.Recordset
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers1_1.ClearField
   Call m_Headers1_2.ClearField
   Call m_Headers1_3.ClearField
   Call m_Headers1_4.ClearField
   Call m_Details1.ClearField
   
   Set Cf = New CReportField
   
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "วัตถุดิบ", "<")
   Call m_Headers0.AddField(Cf)
                                                                                                                                                                                                                                                      
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับเข้า", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "กินสะสม", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "กินเฉลี่ย", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขาย", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เหลือใช้ (วัน)", ">")
   Call m_Headers0.AddField(Cf)
   '===

   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("MONTH1"), "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("MONTH2"), "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("MONTH3"), "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("MONTH4"), "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_2.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "คงเหลือ", "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("LEFT1"), "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("LEFT2"), "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("LEFT3"), "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("LEFT4"), "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_3.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_3.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "ผสมไว้", "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("MIX1"), "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("MIX2"), "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("MIX3"), "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", mcolParam("MIX4"), "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_4.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_4.AddField(Cf)
   
   Call Cf.SetFieldValue(11, "^", "การผสม", "^")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "คลอดสะสม", "^")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "คลอดวันนี้", "^")
   Call m_Headers1_1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "กำหนดคลอด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เป้าหมาย", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เกิดจริง", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ผลต่าง", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เฉลี่ย", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "แม่คลอด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลูกเกิด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@แม่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "~ ลูกเกิด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "แม่คลอด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลูกเกิด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@แม่", "^")
   Call m_Headers1.AddField(Cf)
      
   j = 0
   Set Inv = New CInventoryDoc
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = DateToStringExtEx2(mcolParam("BIRTH_DATE"))
      ElseIf j = 2 Then
         TempStr = mcolParam("TARGET")
      ElseIf j = 3 Then
         TempStr = mcolParam("ACTUAL_BIRTH")
      ElseIf j = 4 Then
         TempStr = mcolParam("DIFF")
      ElseIf j = 5 Then
         TempStr = Format(mcolParam("AVG"), "0.000")
      ElseIf j = 6 Then
         Inv.INVENTORY_DOC_ID = -1
         Inv.DOCUMENT_TYPE = 5
         Inv.FROM_DATE = mcolParam("FROM_DATE")
         Inv.TO_DATE = mcolParam("TO_DATE")
         Inv.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
         Call Inv.QueryData(TempRs, iCount)
         TempStr = iCount
      ElseIf j = 7 Then
         Call LoadPigBirthAmount(Nothing, ImportItems, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
          If ImportItems.Count <= 0 Then
            Amt = 0
          Else
            Amt = ImportItems(1).IMPORT_AMOUNT
          End If
          TempStr = Amt
      ElseIf j = 8 Then
         If iCount <= 0 Then
            TempStr = 0
         Else
            TempStr = FormatNumber(Amt / iCount)
         End If
      ElseIf j = 9 Then
         TempStr = ""
      ElseIf j = 10 Then
         Inv.INVENTORY_DOC_ID = -1
         Inv.DOCUMENT_TYPE = 5
         Inv.FROM_DATE = mcolParam("TO_DATE")
         Inv.TO_DATE = mcolParam("TO_DATE")
         Inv.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
         Call Inv.QueryData(TempRs, iCount)
         TempStr = iCount
      ElseIf j = 11 Then
         Call LoadPigBirthAmount(Nothing, ImportItems, mcolParam("TO_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
          If ImportItems.Count <= 0 Then
            Amt = 0
          Else
            Amt = ImportItems(1).IMPORT_AMOUNT
          End If
          TempStr = Amt
      ElseIf j = 12 Then
         If iCount <= 0 Then
            TempStr = 0
         Else
            TempStr = FormatNumber(Amt / iCount)
         End If
      Else
         TempStr = ""
      End If
      Call Cf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call m_Details1.AddField(Cf)
   Next HeadCf
   Set Inv = Nothing
   
   Set Cf = Nothing
   
  Call m_Headers1_1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
      
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Details1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp = ""
   
  Call m_Headers1_2.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1_3.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1_4.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp = ""
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   Set ImportItems = Nothing
   
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
   ToDate = " ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
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

Private Function GetFirstAvgPrice(Col As Collection, Pi As CPartItem) As Double
Dim Ba As CBalanceAccum
Dim AvgPrice As Double

   AvgPrice = 0
   For Each Ba In Col
      If (Ba.PART_ITEM_ID = Pi.PART_ITEM_ID) And (Ba.MASTER_FLAG = "Y") Then '
         AvgPrice = Ba.AVG_PRICE
         Exit For
      End If
   Next Ba
   
   GetFirstAvgPrice = AvgPrice
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim O As Object
Dim InventoryBals As Collection
Dim RName As String
Dim cData As CPartLocation
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim EI As CExportItem
Dim II As CImportItem
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim TxCode As String
Dim Partitems As Collection
Dim Houses As Collection
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim HS As CLocation
Dim Lc As CLocation
Dim Pi As CPartItem
Dim Ib As CImportItem
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim PiCount As Long
Dim iCount As Long
Dim TempStr As String
Dim Locations As Collection
Dim Sum1 As Double
Dim PrevPI As CPartItem
Dim Yw As CYearWeek
Dim YearWeeks As Collection
Dim PartItemID As Long
Dim NewDate As Date
Dim PigTypeCode As String
Dim BuyItems As Collection
Dim ImportItems As Collection
Dim ImportItemExs As Collection
Dim ExportItems As Collection
Dim ExportItemExs As Collection
Dim PartTypes As Collection
Dim Pt As CPartType
Dim BalanceAmt As Double
Dim ImportAmt As Double
Dim ExportAmt As Double
Dim ImportPrice As Double
Dim UnitPrice As Double
Dim DateCount As Long
Dim AvgEat As Double
Dim LeftAmt As Double
Dim AdvanceAmt As Double
Dim Ba As CBalanceAccum
Dim SumImportPrice As Double
Dim SumImportAmount As Double
Dim SumAllImportPrice As Double
Dim SumAllImportAmount As Double
Dim IntakeFlag As String
Dim ImportPrices As Collection
Dim SumTx As Double
Dim SellItems As Collection
Dim SellAmount As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
         
   DateCount = DateDiff("D", mcolParam("FROM_DATE"), mcolParam("TO_DATE")) + 1
   
   Set ImportItems = New Collection
   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
   Call LoadPartImportAmount(Nothing, ImportItems, -1, NewDate, CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
   
   Set ExportItems = New Collection
   Call LoadPartExportAmount(Nothing, ExportItems, -1, NewDate, CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
      
   Set BuyItems = New Collection
   Call LoadPartImportAmount(Nothing, BuyItems, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), , 1)

   Set ImportItemExs = New Collection
   Call LoadSumBalanceAccum2(Nothing, ImportItemExs, -1, mcolParam("TO_DATE"))

   Set ExportItemExs = New Collection
   Call LoadPartExportAmount(Nothing, ExportItemExs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), 2, "N", mcolParam("HOUSE_GROUP_ID"))

   Set SellItems = New Collection
   Call LoadPartSellAmount(Nothing, SellItems, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))

   Set InventoryBals = New Collection
   'ใช้ TO_DATE แทน FROM_DATE
   Call LoadInventoryBalanceEx(Nothing, InventoryBals, mcolParam("TO_DATE"), mcolParam("TO_DATE"), "")
   
   Set PartTypes = New Collection
   Call LoadPartType(Nothing, PartTypes, mcolParam("PART_GROUP_ID"))
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0

      SumAllImportPrice = 0
      SumAllImportAmount = 0
         
'      If mcolParam("INTAKE_FLAG") = "Y" Then
'         IntakeFlag = ""
'      ElseIf mcolParam("INTAKE_FLAG") = "N" Then
'         IntakeFlag = "Y"
'      End If
   
      For Each Pt In PartTypes
         Set Partitems = Nothing
         Set Partitems = New Collection
         Call LoadImportPartItem(Nothing, Partitems, Pt.PART_TYPE_ID, , mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), "Y")

         PiCount = 0
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
         
         SumImportPrice = 0
         SumImportAmount = 0
         
         For Each Pi In Partitems
'            If (Pi.PART_ITEM_ID = 13268 Or Pi.PART_ITEM_ID = 10443) Then
'               ''debug.print
'            End If
            SumTx = 0
            PiCount = PiCount + 1
            Set BodyCf = New CReportField
            I = I + 1

            Call m_Details1.ClearField
            Sum1 = 0
            
            Set II = GetImportItem(ImportItems, Trim(Str(Pi.PART_ITEM_ID)))
            Set EI = GetExportItem(ExportItems, Trim(Str(Pi.PART_ITEM_ID)))
            BalanceAmt = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT

            Set II = GetImportItem(BuyItems, Trim(Str(Pi.PART_ITEM_ID)))
            ImportAmt = II.IMPORT_AMOUNT
            Set Ba = GetBalanceAccum(ImportItemExs, Trim(Str(Pi.PART_ITEM_ID)))
            ImportPrice = Ba.TOTAL_INCLUDE_PRICE
            UnitPrice = GetFirstAvgPrice(InventoryBals, Pi)
            If UnitPrice <= 0 Then
               UnitPrice = MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
            End If

            SumImportAmount = SumImportAmount + Ba.BALANCE_AMOUNT
            SumAllImportAmount = SumAllImportAmount + Ba.BALANCE_AMOUNT
            SumImportPrice = SumImportPrice + ImportPrice
            SumAllImportPrice = SumAllImportPrice + ImportPrice
            
            Set EI = GetExportItem(ExportItemExs, Trim(Str(Pi.PART_ITEM_ID)))
           ExportAmt = EI.EXPORT_AMOUNT
            
            Set EI = GetExportItem(SellItems, Trim(Str(Pi.PART_ITEM_ID)))
           SellAmount = EI.EXPORT_AMOUNT
            
            If DateCount > 0 Then
               AvgEat = ExportAmt / DateCount
            Else
               AvgEat = 0
            End If
            
            LeftAmt = ImportAmt + BalanceAmt - ExportAmt - SellAmount
            
            If AvgEat > 0 Then
               AdvanceAmt = LeftAmt / AvgEat
            Else
               AdvanceAmt = 0
            End If
            
            TempStr = I
            Set HeadCf = m_Headers0.Fields(1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(1) = Total1(1) + 0
            Total2(1) = Total2(1) + 0
            
            TempStr = Pi.PART_DESC & " (" & Pi.PART_NO & ")"
            Set HeadCf = m_Headers0.Fields(2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(2) = Total1(2) + 0
            Total2(2) = Total2(2) + 0
            
            Set HeadCf = m_Headers0.Fields(3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(UnitPrice, "-"))
            Call m_Details1.AddField(BodyCf)
            Total1(3) = Total1(3) + UnitPrice
            Total2(3) = Total2(3) + UnitPrice
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(4)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(BalanceAmt, "-"))
            Call m_Details1.AddField(BodyCf)
            Total1(4) = Total1(4) + BalanceAmt
            Total2(4) = Total2(4) + BalanceAmt
            SumTx = SumTx + BalanceAmt
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(5)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(ImportAmt, "-"))
            Call m_Details1.AddField(BodyCf)
            Total1(5) = Total1(5) + ImportAmt
            Total2(5) = Total2(5) + ImportAmt
            SumTx = SumTx + ImportAmt
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(6)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(ExportAmt, "-"))
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6) + ExportAmt
            Total2(6) = Total2(6) + ExportAmt
            SumTx = SumTx + ExportAmt
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(7)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(AvgEat, "-"))
            Call m_Details1.AddField(BodyCf)
            Total1(7) = Total1(7) + AvgEat
            Total2(7) = Total2(7) + AvgEat
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(8)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(SellAmount, "-"))
            Call m_Details1.AddField(BodyCf)
            Total1(8) = Total1(8) + SellAmount
            Total2(8) = Total2(8) + SellAmount
            SumTx = SumTx + SellAmount
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(9)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(LeftAmt, "-"))
            Call m_Details1.AddField(BodyCf)
            Total1(9) = Total1(9) + LeftAmt
            Total2(9) = Total2(9) + LeftAmt
            SumTx = SumTx + LeftAmt
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(10)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(AdvanceAmt, "-"))
            Call m_Details1.AddField(BodyCf)
            Total1(10) = Total1(10) + AdvanceAmt
            Total2(10) = Total2(10) + AdvanceAmt
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            If SumTx > 0 Then
               Call Vsp.AddTable(strFormat, "", TempStr2)
            Else
               I = I - 1
               PiCount = PiCount - 1
            End If
         Next Pi
         
         If PiCount > 0 Then
            Call GenerateFooter("รวม", "***" & Pt.PART_TYPE_NAME & "***", "", Total1, m_Details1, SumImportPrice, SumImportAmount)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
         
         Set BodyCf = Nothing
      Next Pt

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, SumAllImportPrice, SumAllImportAmount)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   
      '=====
      For Each Pt In PartTypes
         Set Partitems = Nothing
         Set Partitems = New Collection
         
         Call LoadImportPartItem(Nothing, Partitems, Pt.PART_TYPE_ID, , mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), "N")
         For Each Pi In Partitems
            Set BodyCf = New CReportField

            Call m_Details1.ClearField
            Sum1 = 0
            
            Set II = GetImportItem(ImportItems, Trim(Str(Pi.PART_ITEM_ID)))
            Set EI = GetExportItem(ExportItems, Trim(Str(Pi.PART_ITEM_ID)))
            BalanceAmt = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT
            
            Set II = GetImportItem(BuyItems, Trim(Str(Pi.PART_ITEM_ID)))
            ImportAmt = II.IMPORT_AMOUNT
            Set Ba = GetBalanceAccum(ImportItemExs, Trim(Str(Pi.PART_ITEM_ID)))
            ImportPrice = Ba.TOTAL_INCLUDE_PRICE
            UnitPrice = GetFirstAvgPrice(InventoryBals, Pi)
            If UnitPrice <= 0 Then
               UnitPrice = MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
            End If
            
            Set EI = GetExportItem(ExportItemExs, Trim(Str(Pi.PART_ITEM_ID)))
           ExportAmt = EI.EXPORT_AMOUNT
            
            Set EI = GetExportItem(SellItems, Trim(Str(Pi.PART_ITEM_ID)))
           SellAmount = EI.EXPORT_AMOUNT
           
            If DateCount > 0 Then
               AvgEat = ExportAmt / DateCount
            Else
               AvgEat = 0
            End If
            
            LeftAmt = ImportAmt + BalanceAmt - ExportAmt - SellAmount
            
            If AvgEat > 0 Then
               AdvanceAmt = LeftAmt / AvgEat
            Else
               AdvanceAmt = 0
            End If
            
            TempStr = "***"
            Set HeadCf = m_Headers0.Fields(1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
'            Total1(1) = Total1(1) + 0
'            Total2(1) = Total2(1) + 0
            
            TempStr = Pi.PART_DESC & " (" & Pi.PART_NO & ")"
            Set HeadCf = m_Headers0.Fields(2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
'            Total1(2) = Total1(2) + 0
'            Total2(2) = Total2(2) + 0
            
            Set HeadCf = m_Headers0.Fields(3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(UnitPrice, "-"))
            Call m_Details1.AddField(BodyCf)
'            Total1(3) = Total1(3) + UnitPrice
'            Total2(3) = Total2(3) + UnitPrice
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(4)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(BalanceAmt, "-"))
            Call m_Details1.AddField(BodyCf)
'            Total1(4) = Total1(4) + BalanceAmt
'            Total2(4) = Total2(4) + BalanceAmt
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(5)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(ImportAmt, "-"))
            Call m_Details1.AddField(BodyCf)
'            Total1(5) = Total1(5) + ImportAmt
'            Total2(5) = Total2(5) + ImportAmt
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(6)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(ExportAmt, "-"))
            Call m_Details1.AddField(BodyCf)
'            Total1(6) = Total1(6) + ExportAmt
'            Total2(6) = Total2(6) + ExportAmt
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(7)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(AvgEat, "-"))
            Call m_Details1.AddField(BodyCf)
'            Total1(7) = Total1(7) + AvgEat
'            Total2(7) = Total2(7) + AvgEat
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(8)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(SellAmount, "-"))
            Call m_Details1.AddField(BodyCf)
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(9)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(LeftAmt, "-"))
            Call m_Details1.AddField(BodyCf)
'            Total1(9) = Total1(9) + LeftAmt
'            Total2(9) = Total2(9) + LeftAmt
            
            TempStr = 0
            Set HeadCf = m_Headers0.Fields(10)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, FormatNumber(AdvanceAmt, "-"))
            Call m_Details1.AddField(BodyCf)
'            Total1(10) = Total1(10) + AdvanceAmt
'            Total2(10) = Total2(10) + AdvanceAmt
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         Next Pi
      Next Pt
      '=====
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set Rs2 = Nothing
         
   Set YearWeeks = Nothing
   Set Partitems = Nothing
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   Set PartTypes = Nothing
   Set ImportItemExs = Nothing
   Set InventoryBals = Nothing
   Set SellItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, SumImportPrice As Double, SumImportAmount As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = FormatNumber(MyDiff(SumImportPrice, SumImportAmount))
      ElseIf (j = 7) Or (j = 10) Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(j))
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
   Vsp.Paragraph = ""
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub


