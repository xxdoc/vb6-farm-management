VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory029"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory029"
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
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

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
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ImportBals As Collection
Private m_ExportBals As Collection


Private PigSumEveryDay As Collection
Private PigSumPeriodDay As Collection

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
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
   Set m_ImportBals = New Collection
   Set m_ExportBals = New Collection
   
   Set PigSumEveryDay = New Collection
   Set PigSumPeriodDay = New Collection
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
   Set m_Details2 = Nothing
   Set m_ImportBals = Nothing
   Set m_ExportBals = Nothing
   
   Set PigSumEveryDay = Nothing
   Set PigSumPeriodDay = Nothing
   
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
   
   Set Cf = New CReportField
   
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "สป. เกิด", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "อายุ", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "สุกร", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", mcolParam("STATUS_GROUP_NAME1"), "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", mcolParam("STATUS_GROUP_NAME2"), "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", mcolParam("PART_GROUP_NAME"), "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "intake", "<")
   Call m_Headers0.AddField(Cf)
   '===
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "สป. เกิด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "อายุ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เฉลี่ย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "สะสม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เฉลี่ย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันนี้", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "สะสม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เฉลี่ย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันนี้", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "สะสม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เฉลี่ย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันนี้", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "intake", ">")
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
                                       vbCrLf & FromDate & ToDate & _
                                       vbCrLf & "กลุ่มโรงเรือน " & EmptyToString(mcolParam("HOUSE_GROUP_NAME"), "N/A")
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
Dim Total3(100) As Double
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
Dim DeadPigs As Collection
Dim TodayDeadPigs As Collection
Dim DeadAmt As Double
Dim AvgPig As Double
Dim PigAmt As Double
Dim RejectPigs As Collection
Dim TodayRejectPigs As Collection
Dim RejectAmt As Double
Dim PigUsed As Collection
Dim TodayPigUsed As Collection
Dim AvgUsed As Double
Dim TempBa As CBalanceAccum
   
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   DateCount = DateDiff("D", mcolParam("FROM_DATE"), mcolParam("TO_DATE")) + 1
   
   Set ImportItems = New Collection            'คิดว่าน่าจะตัดออกได้โดยเอาเฉพาะหมูที่ยังคงคำนวณอยู่
   Call LoadPigImportAmount(Nothing, ImportItems, -1, mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), , , , , mcolParam("HOUSE_GROUP_ID"))
   
   Set ExportItems = New Collection
   Call LoadPigExportAmount(Nothing, ExportItems, -1, mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), , , , , , mcolParam("HOUSE_GROUP_ID"))
   
'   Set ImportItemExs = New Collection
'   Call LoadPigImportAmount(Nothing, ImportItemExs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
   
'   Set ExportItemExs = New Collection
'   Call LoadPigExportAmount(Nothing, ExportItemExs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
   
   Set DeadPigs = New Collection
   If mcolParam("MODE") = 1 Then 'คิดตามเอกสารการโอน
      Call LoadPigStatusAmount(Nothing, DeadPigs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID1"), , 7, mcolParam("HOUSE_GROUP_ID"))
   ElseIf mcolParam("MODE") = 2 Then 'คิดตามบิลขาย
      Call LoadPigStatusAmount(Nothing, DeadPigs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID1"), , , mcolParam("HOUSE_GROUP_ID"), "(10,13)")
   End If
   
   Set TodayDeadPigs = New Collection
   If mcolParam("MODE") = 1 Then 'คิดตามเอกสารการโอน
      Call LoadPigStatusAmount(Nothing, TodayDeadPigs, mcolParam("TO_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID1"), , 7, mcolParam("HOUSE_GROUP_ID"))
   ElseIf mcolParam("MODE") = 2 Then 'คิดตามบิลขาย
      Call LoadPigStatusAmount(Nothing, TodayDeadPigs, mcolParam("TO_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID1"), , , mcolParam("HOUSE_GROUP_ID"), "(10,13)")
   End If
   
   Set RejectPigs = New Collection
   If mcolParam("MODE") = 1 Then 'คิดตามเอกสารการโอน
      Call LoadPigStatusAmount(Nothing, RejectPigs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID2"), , 7, mcolParam("HOUSE_GROUP_ID"))
   ElseIf mcolParam("MODE") = 2 Then 'คิดตามบิลขาย
      Call LoadPigStatusAmount(Nothing, RejectPigs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID2"), , , mcolParam("HOUSE_GROUP_ID"), "(10,13)")
   End If
   
   Set TodayRejectPigs = New Collection
   If mcolParam("MODE") = 1 Then 'คิดตามเอกสารการโอน
      Call LoadPigStatusAmount(Nothing, TodayRejectPigs, mcolParam("TO_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID2"), , 7, mcolParam("HOUSE_GROUP_ID"))
   ElseIf mcolParam("MODE") = 2 Then 'คิดตามบิลขาย
      Call LoadPigStatusAmount(Nothing, TodayRejectPigs, mcolParam("TO_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID2"), , , mcolParam("HOUSE_GROUP_ID"), "(10,13)")
   End If
   
   Set Partitems = New Collection
   Call LoadImportPigEx(Nothing, Partitems, , mcolParam("TO_DATE"))
   Call Selectionsort(Partitems, 1, Partitems.Count, 1)
   
   Set PigUsed = New Collection
   Call LoadPigPartUseAmount(Nothing, PigUsed, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), "Y", mcolParam("HOUSE_GROUP_ID"))
   
   Set TodayPigUsed = New Collection
   Call LoadPigPartUseAmount(Nothing, TodayPigUsed, mcolParam("TO_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), "Y", mcolParam("HOUSE_GROUP_ID"))
   
   Call GenerateBalanceAccumSumEveryDay(PigSumEveryDay, -1, DateAdd("D", -1, mcolParam("FROM_DATE")), mcolParam("HOUSE_GROUP_ID"))
   Call GenerateBalanceAccumSumDate(PigSumPeriodDay, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("HOUSE_GROUP_ID"))
   
   If Not initDoc Then
      Exit Function
   End If
   
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
      Set BodyCf = New CReportField
      
      For j = 1 To UBound(Total1)
         Total3(j) = 0
      Next j
      
      Set PrevPI = New CPartItem
      For Each Pi In Partitems
         I = I + 1
         
'         If Pi.PART_ITEM_ID = 16554 Then
'            Debug.Print
'         End If
         
         Call m_Details1.ClearField
         Set EI = GetExportItem(ExportItems, Trim(Str(Pi.PART_ITEM_ID)))
         Set II = GetImportItem(ImportItems, Trim(Str(Pi.PART_ITEM_ID)))
         LeftAmt = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT
                  
         '===
If Pi.AGE_CODE = "R" Then
'''debug.print
End If
         If PrevPI.AGE_CODE <> Pi.AGE_CODE Then
            If (I > 1) Then
               Call GenerateFooter("รวม", PrevPI.AGE_CODE, "", Total1, m_Details2, DateCount)
               Call m_Details2.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               Set PrevPI = Pi
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
            Else
               Set PrevPI = Pi
            End If
         End If
         '===
         
         j = 0
         For Each HeadCf In m_Headers1.Fields
            j = j + 1
            If j = 1 Then
               TempStr = I
            ElseIf j = 2 Then
If Pi.PART_ITEM_ID = 12783 Then
'''debug.print Pi.PART_ITEM_ID
End If
               TempStr = Pi.PART_NO & " (" & Pi.PIG_TYPE & ")"
            ElseIf j = 3 Then
               TempStr = Pi.PIG_AGE
'               If Pi.PIG_AGE > 800 Then
'                  Debug.Print
'               End If
            ElseIf j = 4 Then
'               Amt = GenerateBalanceAmtLocal(Pi.PART_ITEM_ID, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
               
               Amt = GenerateBalanceAmt(Pi.PART_ITEM_ID, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("HOUSE_GROUP_ID"))             'ยกเลิกใช้เนื่องจาก Query ตลอดเวลา
               PigAmt = Amt
               
               TempStr = FormatNumber(Amt, , 0)
            ElseIf j = 5 Then
               If DateCount > 0 Then
                  Amt = Amt / DateCount 'Use previous value j =4
               Else
                  Amt = 0
               End If
               AvgPig = Amt
               TempStr = FormatNumber(Amt)
            ElseIf j = 6 Then
               Amt = LeftAmt
               TempStr = FormatNumber(Amt, , 0)
            ElseIf j = 7 Then
               Set EI = GetExportItem(DeadPigs, Trim(Str(Pi.PART_ITEM_ID)))
               Amt = EI.EXPORT_AMOUNT
               DeadAmt = Amt
               
               TempStr = FormatNumber(Amt, , 0)
            ElseIf j = 8 Then
               If AvgPig <> 0 Then
                  Amt = (DeadAmt / AvgPig) * 100
               Else
                  Amt = 0
               End If
               
               TempStr = FormatNumber(Amt)
            ElseIf j = 9 Then
               If DateCount <> 0 Then
                  Amt = (DeadAmt / DateCount)
               Else
                  Amt = 0
               End If
               
               TempStr = FormatNumber(Amt)
            ElseIf j = 10 Then
               Set EI = GetExportItem(TodayDeadPigs, Trim(Str(Pi.PART_ITEM_ID)))
               Amt = EI.EXPORT_AMOUNT
               
               TempStr = FormatNumber(Amt, , 0)
            ElseIf j = 11 Then
               Set EI = GetExportItem(RejectPigs, Trim(Str(Pi.PART_ITEM_ID)))
               Amt = EI.EXPORT_AMOUNT
               RejectAmt = Amt
               
               TempStr = FormatNumber(Amt, , 0)
            ElseIf j = 12 Then
               If AvgPig <> 0 Then
                  Amt = (RejectAmt / AvgPig) * 100
               Else
                  Amt = 0
               End If
               
               TempStr = FormatNumber(Amt)
            ElseIf j = 13 Then
               If DateCount <> 0 Then
                  Amt = (RejectAmt / DateCount)
               Else
                  Amt = 0
               End If
               
               TempStr = FormatNumber(Amt)
            ElseIf j = 14 Then
               Set EI = GetExportItem(TodayRejectPigs, Trim(Str(Pi.PART_ITEM_ID)))
               Amt = EI.EXPORT_AMOUNT
               
               TempStr = FormatNumber(Amt, , 0)
            ElseIf j = 15 Then
               Set EI = GetExportItem(PigUsed, Pi.PART_ITEM_ID & "-" & mcolParam("PART_GROUP_ID") & "-N")
               Amt = EI.EXPORT_AMOUNT
               
               TempStr = FormatNumber(Amt)
            ElseIf j = 16 Then
               If DateCount <> 0 Then
                  Amt = Amt / DateCount
               Else
                  Amt = 0
               End If
               AvgUsed = Amt
               
               TempStr = FormatNumber(Amt)
            ElseIf j = 17 Then
               Set EI = GetExportItem(TodayPigUsed, Pi.PART_ITEM_ID & "-" & mcolParam("PART_GROUP_ID") & "-N")
               Amt = EI.EXPORT_AMOUNT
               
               TempStr = FormatNumber(Amt)
            ElseIf j = 18 Then
               If AvgPig <> 0 Then
                  Amt = AvgUsed / AvgPig
               Else
                  Amt = 0
               End If
               TempStr = FormatNumber(Amt)
            Else
               TempStr = ""
            End If
            
            If LeftAmt + PigAmt <> 0 Then
               Total1(j) = Total1(j) + Amt
               Total2(j) = Total2(j) + Amt
            End If
            
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         Next HeadCf
      
         If LeftAmt + PigAmt <> 0 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         Else
            I = I - 1
         End If
      Next Pi
      
      If I > 1 Then
         Call GenerateFooter("รวม", PrevPI.AGE_CODE, "", Total1, m_Details2, DateCount)
         Call m_Details2.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Set PrevPI = Pi
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
      Else
         Set PrevPI = Pi
      End If
      
      If I > 1 Then
         Call GenerateFooter("สุทธิ", "", "", Total2, m_Details2, DateCount)
         Call m_Details2.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      
      Set PrevPI = Nothing
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set Rs2 = Nothing
         
   Set YearWeeks = Nothing
   Set Partitems = Nothing
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   Set PartTypes = Nothing
   Set ImportItemExs = Nothing
   Set DeadPigs = Nothing
   Set TodayDeadPigs = Nothing
   Set RejectPigs = Nothing
   Set TodayRejectPigs = Nothing
   Set PigUsed = Nothing
   Set TodayPigUsed = Nothing
   
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
      ElseIf (j = 3) Then
         TempStr = ""
      ElseIf j = 5 Then
         TempStr = FormatNumber(MyDiffEx(Tot(4), DateCount))
      ElseIf j = 8 Then
         TempStr = FormatNumber(MyDiffEx(Tot(7), Tot(5)) * 100)
      ElseIf j = 9 Then
         TempStr = FormatNumber(MyDiffEx(Tot(7), DateCount))
      ElseIf j = 12 Then
         TempStr = FormatNumber(MyDiffEx(Tot(11), Tot(5)) * 100)
      ElseIf j = 13 Then
         TempStr = FormatNumber(MyDiffEx(Tot(11), DateCount))
      ElseIf j = 15 Then
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 16 Then
         TempStr = FormatNumber(MyDiffEx(Tot(15), DateCount))
      ElseIf j = 17 Then
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 18 Then
         TempStr = FormatNumber(MyDiffEx(Tot(16), Tot(5)))
      Else
         TempStr = FormatNumber(Tot(j), , 0)
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
Private Function GenerateBalanceAmtLocal(PartItemID As Long, FromDate As Date, ToDate As Date) As Double
Dim TempSum As Double
Dim Ba As CBalanceAccum
Dim NewDate As Date
Dim DaySum As Double
   Set Ba = New CBalanceAccum
   
   NewDate = FromDate 'mcolParam("FROM_DATE")
   
'   If (PartItemID = 17318) Then
'      Debug.Print
'   End If
   
   Set Ba = GetBalanceAccum(PigSumEveryDay, Trim(Str(PartItemID)))
   
   DaySum = Ba.BALANCE_AMOUNT
   While NewDate <= ToDate 'mcolParam("TO_DATE")
      Set Ba = GetBalanceAccum(PigSumPeriodDay, Trim(PartItemID & "-" & NewDate))
      
      If (Ba.BALANCE_AMOUNT > 0) Then
         Debug.Print
      End If
      
      DaySum = DaySum + Ba.BALANCE_AMOUNT
      TempSum = TempSum + DaySum
      
      NewDate = DateAdd("D", 1, NewDate)
   Wend
   
   Set Ba = Nothing
   
   GenerateBalanceAmtLocal = TempSum
End Function

