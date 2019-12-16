VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSell008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSell008"
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

Private m_SellPrices As Collection
Private m_Customers As Collection
Private m_PartItems As Collection
Private Cl As Collection
Private CL2 As Collection


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
   Set m_SellPrices = New Collection
   Set m_Customers = New Collection
   Set m_PartItems = New Collection
   Set Cl = New Collection
      Set CL2 = New Collection
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
   Set m_SellPrices = Nothing
   Set m_Customers = Nothing
   Set m_PartItems = Nothing
   Set Cl = Nothing
      Set CL2 = Nothing
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
   Dim i As Long
   
   For i = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next i

End Sub
'============================= Codes above used in every report ==========================

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(7, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัสลูกค้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "ประเภท", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ต้นทุน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา/หน่วย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "GP", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%GP", ">")
   Call m_Headers1.AddField(Cf)
   Set Cf = Nothing

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
Dim ToPaidDate As String
Dim NewDate As Date

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadSalePartCustAmountPrice_SLMKEY(Nothing, m_SellPrices, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", mcolParam("CUSTOMER_TYPE"))
   Call LoadPartItem(Nothing, m_PartItems, , "N")
   Call LoadCustomer(Nothing, m_Customers)

   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")

   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate
      
      If Len(mcolParam("FROM_SALE_CODE")) > 0 Then
              MODULE_DESC = MODULE_DESC & "         พนักงานขาย  " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A")
      End If
       If Len(mcolParam("TO_SALE_CODE")) > 0 Then
              MODULE_DESC = MODULE_DESC & "   ถึง    " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
      End If
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim i As Long
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
Dim Total4(100) As Double
Dim Total5(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim TempApt As CDoItem
Dim Hc As CReportField
Dim Di1 As CDoItem
Dim temp_Di1 As CDoItem
Dim Di2 As CDoItem
Dim Pi As CPartItem
Dim Cm As CCustomer
Dim PrevPartDesc As String
Dim PrevPartNo As String
Dim PrevPartType As String
Dim PrevPartID As Long
Dim PrevEmpName As String
Dim PrevEmpCode As String
Dim PrevCustType As Long
Dim PrevCustTypeName As String
Dim SubHeaderFlag As Boolean
Dim TotSum As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
      Total4(j) = 0
      Total5(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
       Set BodyCf = New CReportField
      Set Di1 = New CDoItem
      Di1.DO_ITEM_ID = -1
      Di1.PIG_FLAG = "N"
      Di1.FROM_DATE = mcolParam("FROM_DATE")
      Di1.TO_DATE = mcolParam("TO_DATE")
      Di1.FROM_EMP_CODE = mcolParam("FROM_SALE_CODE")
      Di1.TO_EMP_CODE = mcolParam("TO_SALE_CODE")
      Di1.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Di1.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Di1.OrderBy = 3
      Call Di1.QueryData(9, Rs, iCount)
      
      i = 0
      SubHeaderFlag = False
      While Not Rs.EOF
         i = i + 1
         Call Di1.PopulateFromRS(9, Rs)
              
'       Set temp_Di1 = Nothing
'      Set temp_Di1 = GetObject("CDoItem", CL, Trim(Str(Di1.DO_ITEM_ID)), False)
'      If temp_Di1 Is Nothing Then
'           Call CL.Add(Di1, Trim(Str(Di1.DO_ITEM_ID)))
'      Else
'           ''debug.print Di1.DO_ID
'           Call CL2.Add(Di1, Trim(Str(Di1.DO_ITEM_ID)))
'      End If
         
         Set Pi = GetPartItem(m_PartItems, Di1.PART_ITEM_ID)
         Set Cm = GetCustomer(m_Customers, Di1.CUSTOMER_ID)
         Set Di2 = GetDoItem(m_SellPrices, Di1.CUSTOMER_ID & "-" & Di1.PART_ITEM_ID & "-" & Di1.EMP_ID)
         
         If Di1.EMP_ID <> 0 Then
            ''debug.print
         End If
         
         
         If i = 1 Then
            PrevCustType = Cm.CUSTOMER_TYPE
            PrevPartID = Pi.PART_ITEM_ID
            PrevPartDesc = Pi.PART_DESC
            PrevPartNo = Pi.PART_NO
            PrevPartType = Pi.PART_TYPE_NAME
            PrevEmpName = Di1.EMP_LONGNAME
            PrevEmpCode = Di1.EMP_CODE
            PrevCustTypeName = Cm.CSTTYPE_NAME
         
            Vsp = "พนักงานขาย : (" & PrevEmpCode & ")    " & PrevEmpName
         
           If mcolParam("PART_TYPE") = "N" Then
                Call GenerateSubHeader(PrevPartType, "", m_Details1, 0)    ' ประเภทสินค้า
                Call m_Details1.GetString(1, TempStr1, TempStr2)
                strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                Vsp.FontBold = True
                Call Vsp.AddTable(strFormat, "", TempStr2)
                Vsp.FontBold = False
            End If
         
            If mcolParam("NOT_PART_FLAG") = "N" And mcolParam("NOT_CUS_FLAG") = "N" Then
                Call GenerateSubHeader(PrevPartNo, PrevPartDesc, m_Details1, 0)   ' & "    (" & PrevPartType & ")"
                Call m_Details1.GetString(1, TempStr1, TempStr2)
                strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                Vsp.FontBold = True
                Call Vsp.AddTable(strFormat, "", TempStr2)
                Vsp.FontBold = False
             End If
   
         Else
         
            If PrevPartID <> Pi.PART_ITEM_ID And mcolParam("NOT_PART_FLAG") = "N" Then
               'รวมตามลูกค้า
               Call GenerateFooter("รวม", PrevPartNo, PrevPartDesc, "", Total2, m_Details1, 0)               ' PrevPartType
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
              If mcolParam("NOT_CUS_FLAG") = "N" Then
                  Vsp.FontBold = True
               End If
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            
               For j = 1 To UBound(Total1)
                  Total2(j) = 0
               Next j
            PrevPartID = Pi.PART_ITEM_ID
            PrevPartDesc = Pi.PART_DESC
            PrevPartNo = Pi.PART_NO
            
               SubHeaderFlag = True
            End If
            
          If PrevPartType <> Pi.PART_TYPE_NAME Then
               'รวมตามลูกค้า
             If mcolParam("PART_TYPE") = "N" Then
               Call GenerateFooter("", "รวม", PrevPartType, "", Total5, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
               
               For j = 1 To UBound(Total5)
                  Total5(j) = 0
               Next j
         
           End If
            
             If PrevCustType <> Cm.CUSTOMER_TYPE Then
               'รวมตามประเภทลูกค้า
               Call GenerateFooter("รวม", PrevCustTypeName, "", "", Total3, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            
               For j = 1 To UBound(Total1)
                  Total3(j) = 0
               Next j
               PrevCustType = Cm.CUSTOMER_TYPE
               PrevCustTypeName = Cm.CSTTYPE_NAME
            End If
            
            If PrevEmpCode <> Di1.EMP_CODE Then
               'รวมตามพนักงาน
               Call GenerateFooterSale("รวมพนักงาน  :  (" & PrevEmpCode & ")  " & PrevEmpName, Total4, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            
               For j = 1 To UBound(Total1)
                  Total4(j) = 0
               Next j
               PrevEmpName = Di1.EMP_LONGNAME
               PrevEmpCode = Di1.EMP_CODE
            
               Vsp = "พนักงานขาย : (" & PrevEmpCode & ")    " & PrevEmpName
               SubHeaderFlag = True
            End If
         End If
         
           If PrevPartType <> Pi.PART_TYPE_NAME Then
               If mcolParam("PART_TYPE") = "N" Then
                  Call GenerateSubHeader(Pi.PART_TYPE_NAME, "", m_Details1, 0)     ' ประเภทสินค้า
                   Call m_Details1.GetString(1, TempStr1, TempStr2)
                   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                   Vsp.FontBold = True
                   Call Vsp.AddTable(strFormat, "", TempStr2)
                   Vsp.FontBold = False
                End If
               PrevPartType = Pi.PART_TYPE_NAME
            End If

         If SubHeaderFlag Then
            'New header
            
             If mcolParam("NOT_PART_FLAG") = "N" And mcolParam("NOT_CUS_FLAG") = "N" Then
               Call GenerateSubHeader(PrevPartNo, PrevPartDesc, m_Details1, 0)               ' & "    (" & PrevPartType & ")"
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
             End If
            
            SubHeaderFlag = False
         End If
         
         Call m_Details1.ClearField
          Set Hc = m_Headers1.Fields(1)
         TempStr = i
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          Set Hc = m_Headers1.Fields(2)
         TempStr = Cm.CUSTOMER_CODE
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          Set Hc = m_Headers1.Fields(3)
         TempStr = Cm.CUSTOMER_NAME
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          Set Hc = m_Headers1.Fields(4)
         TempStr = ""
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          Set Hc = m_Headers1.Fields(5)
         TempStr = FormatNumber(Di2.EXPORT_TOTAL_PRICE)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + (Di2.EXPORT_TOTAL_PRICE)
         Total2(5) = Total2(5) + (Di2.EXPORT_TOTAL_PRICE)
         Total3(5) = Total3(5) + (Di2.EXPORT_TOTAL_PRICE)
         Total4(5) = Total4(5) + (Di2.EXPORT_TOTAL_PRICE)
         Total5(5) = Total5(5) + (Di2.EXPORT_TOTAL_PRICE)
         
          Set Hc = m_Headers1.Fields(6)
         TempStr = FormatNumber(Di2.ITEM_AMOUNT)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + (Di2.ITEM_AMOUNT)
         Total2(6) = Total2(6) + (Di2.ITEM_AMOUNT)
         Total3(6) = Total3(6) + (Di2.ITEM_AMOUNT)
         Total4(6) = Total4(6) + (Di2.ITEM_AMOUNT)
         Total5(6) = Total5(6) + (Di2.ITEM_AMOUNT)

          Set Hc = m_Headers1.Fields(7)
         TempStr = FormatNumber(Di2.TOTAL_PRICE)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + (Di2.TOTAL_PRICE)
         Total2(7) = Total2(7) + (Di2.TOTAL_PRICE)
         Total3(7) = Total3(7) + (Di2.TOTAL_PRICE)
         Total4(7) = Total4(7) + (Di2.TOTAL_PRICE)
         Total5(7) = Total5(7) + (Di2.TOTAL_PRICE)

          Set Hc = m_Headers1.Fields(8)
         TempStr = FormatNumber(MyDiff(Di2.TOTAL_PRICE, Di2.ITEM_AMOUNT))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(8) = Total1(8) + (MyDiff(Di2.TOTAL_PRICE, Di2.ITEM_AMOUNT))
         Total2(8) = Total2(8) + (MyDiff(Di2.TOTAL_PRICE, Di2.ITEM_AMOUNT))
         Total3(8) = Total3(8) + (MyDiff(Di2.TOTAL_PRICE, Di2.ITEM_AMOUNT))
         Total4(8) = Total4(8) + (MyDiff(Di2.TOTAL_PRICE, Di2.ITEM_AMOUNT))
         Total5(8) = Total5(8) + (MyDiff(Di2.TOTAL_PRICE, Di2.ITEM_AMOUNT))

          Set Hc = m_Headers1.Fields(9)
         TempStr = FormatNumber(Di2.TOTAL_PRICE - Di2.EXPORT_TOTAL_PRICE)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + (Di2.TOTAL_PRICE - Di2.EXPORT_TOTAL_PRICE)
         Total2(9) = Total2(9) + (Di2.TOTAL_PRICE - Di2.EXPORT_TOTAL_PRICE)
         Total3(9) = Total3(9) + (Di2.TOTAL_PRICE - Di2.EXPORT_TOTAL_PRICE)
         Total4(9) = Total4(9) + (Di2.TOTAL_PRICE - Di2.EXPORT_TOTAL_PRICE)
         Total5(9) = Total5(9) + (Di2.TOTAL_PRICE - Di2.EXPORT_TOTAL_PRICE)
         
         Set Hc = m_Headers1.Fields(10)
          Amt = MyDiffEx(Di2.TOTAL_PRICE - Di2.EXPORT_TOTAL_PRICE, Di2.TOTAL_PRICE) * 100
         TempStr = FormatNumber(Amt)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(10) = Total1(10) + Amt
         Total2(10) = Total2(10) + Amt
         Total3(10) = Total3(10) + Amt
         Total4(10) = Total4(10) + Amt
         Total5(10) = Total5(10) + Amt
         
      If mcolParam("NOT_CUS_FLAG") = "N" Then
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
       End If
         
         Rs.MoveNext
      Wend

      '===
      'รวมตามลูกค้า
      If mcolParam("NOT_PART_FLAG") = "N" Then
         Call GenerateFooter("รวม", PrevPartNo, PrevPartDesc, "", Total2, m_Details1, 0)        'PrevPartType
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
              If mcolParam("NOT_CUS_FLAG") = "N" Then
                  Vsp.FontBold = True
               End If
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      

               'รวมตามลูกค้า
              If mcolParam("PART_TYPE") = "N" Then
               Call GenerateFooter("", "รวม", PrevPartType, "", Total5, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               End If
      
     Call GenerateFooter("รวม", PrevCustTypeName, "", "", Total3, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
        'รวมตามพนักงาน
         Call GenerateFooterSale("รวมพนักงาน  :  (" & PrevEmpCode & ")  " & PrevEmpName, Total4, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False

      Call GenerateFooter("สุทธิ", "", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False


      Vsp = ""

      'หาค่า เพิ่ม/ลดหนี้
      Dim Bd  As CBillingDoc
      Set Bd = New CBillingDoc
      Bd.BILLING_DOC_ID = -1
      Bd.FROM_DATE = mcolParam("FROM_DATE")
      Bd.TO_DATE = mcolParam("TO_DATE")
      Bd.ItemSumFlag = True
      Bd.DocType34 = True
      Call Bd.SetFlag(True, True, False, False, False, False)
      Call glbDaily.QueryBillingDoc(Bd, Rs, iCount, IsOK, glbErrorLog)
      While Not Rs.EOF
         Call Bd.PopulateFromRS(1, Rs)
            If (Bd.DOCUMENT_TYPE = 3) Or (Bd.DOCUMENT_TYPE = 4) Then  'เพิ่มหนี้, ลดหนี้
               Call GenerateReceiptItem(Bd, m_Details1, TotSum, Bd.DOCUMENT_TYPE)
            End If
         Rs.MoveNext
      Wend

     Call GenerateFooterReceipt("เพิ่ม/ลดหนี้", TotSum, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooterReceipt("สุทธิ", Total1(9) + TotSum, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False





      Set Di1 = Nothing
      Set BodyCf = Nothing
    Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
''   ''debug.print CL2.Count
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateSubHeader(TXT1 As String, TXT2 As String, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   TempStr = TXT1
   TempWidth = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)

   TempStr = TXT2
   TempWidth = 0
   For j = 3 To m_Headers1.Fields.Count
      TempWidth = TempWidth + m_Headers1.Fields(j).Width
   Next j
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "^")
   Call Details.AddField(BodyCf)

   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooterSale(TXT1 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
         TempStr = TXT1
      ElseIf j = 8 Then
         TempStr = FormatNumber(MyDiffEx(Tot(7), Tot(6)))
      ElseIf j = 9 Then
         TempStr = FormatNumber(Tot(7) - Tot(5))
      ElseIf j = 10 Then
         TempStr = FormatNumber(MyDiff((Tot(7) - Tot(5)), Tot(7) * 100))
      Else
         TempStr = FormatNumber(Tot(j + 3))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Txt4 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
         TempStr = TXT1
      ElseIf j = 2 Then
         TempStr = TXT2
      ElseIf j = 3 Then
         TempStr = TXT3
      ElseIf j = 4 Then
         TempStr = ""    'Txt4
      ElseIf j = 8 Then
         TempStr = FormatNumber(MyDiffEx(Tot(7), Tot(6)))
      ElseIf j = 9 Then
         TempStr = FormatNumber(Tot(7) - Tot(5))
      ElseIf j = 10 Then
         TempStr = FormatNumber(MyDiff((Tot(7) - Tot(5)), Tot(7) * 100))
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
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(Now), 8))
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

Private Sub GenerateReceiptItem(TempBD As CBillingDoc, Details As CFieldList, Sum As Double, DocumentType As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim Bd As CBillingDoc
Dim alngX() As Long
Dim j As Long
Dim Ri As CReceiptItem
Dim TempStr1 As String
Dim TempStr2 As String
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim strFormat As String
Dim Tot() As Double
Dim Ri1_0 As CReceiptItem

   Set TempRs = New ADODB.Recordset
   
   Set Bd = New CBillingDoc
   Bd.QueryFlag = 1
   Bd.BILLING_DOC_ID = TempBD.BILLING_DOC_ID
   Call glbDaily.QueryBillingDoc(Bd, TempRs, iCount, IsOK, glbErrorLog)

     For Each Ri In Bd.ReceiptItems
            If DocumentType = 3 Then
               TempStr1 = FormatNumber(Ri.DEBIT_CREDIT_AMOUNT, "-")
               Sum = Sum + (Ri.DEBIT_CREDIT_AMOUNT)
            ElseIf DocumentType = 4 Then
               TempStr1 = FormatNumber(-1 * Ri.DEBIT_CREDIT_AMOUNT, "-")
               Sum = Sum + (-1 * Ri.DEBIT_CREDIT_AMOUNT)
            End If
   Next Ri
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set BodyCf = Nothing
   Set Bd = Nothing
End Sub

Private Sub GenerateFooterReceipt(TXT1 As String, Sum As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 2 Then
         TempStr = TXT1
      ElseIf j = 9 Then
         TempStr = FormatNumber(Sum)
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

