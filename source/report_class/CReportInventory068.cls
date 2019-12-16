VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory068"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory062"
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
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

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

Private m_SumGroup1 As Collection
Private m_SumGroup2 As Collection
Private m_SumGroup3 As Collection

Private m_SumGroupNew1 As Collection
Private m_SumGroupNew2 As Collection
Private m_SumGroupNew3 As Collection

Private m_SumEx1 As Collection
Private m_SumEx2 As Collection
Private m_SumEx3 As Collection
Private m_SumEx4 As Collection

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
   Set m_SumGroup1 = New Collection
   Set m_SumGroup2 = New Collection
   Set m_SumGroup3 = New Collection
   
   Set m_SumGroupNew1 = New Collection
   Set m_SumGroupNew2 = New Collection
   Set m_SumGroupNew3 = New Collection
   
   Set m_SumEx1 = New Collection
   Set m_SumEx2 = New Collection
   Set m_SumEx3 = New Collection
   Set m_SumEx4 = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_SumGroup1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_SumGroup2 = Nothing
   Set m_SumGroup3 = Nothing
   
   Set m_SumGroupNew1 = Nothing
   Set m_SumGroupNew2 = Nothing
   Set m_SumGroupNew3 = Nothing
   
   Set m_SumEx1 = Nothing
   Set m_SumEx2 = Nothing
   Set m_SumEx3 = Nothing
   Set m_SumEx4 = Nothing
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
Dim ExportWidth As Double
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempFromWeight As Double
Dim TempToWeight As Double
    
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

    
   Call Cf.SetFieldValue(5, "^", "ลูกค้า", "<")
    Call m_Headers1.AddField(Cf)
    
   Call Cf.SetFieldValue(5, "^", "***", ">")
   Call m_Headers0.AddField(Cf)
   
   
    'จำนวน
    '------------------------------------------------------------------------------------------------------------------------------
    ExportWidth = 0
    If Val(mcolParam("LESS_WEIGHT")) > 0 Then
        Call Cf.SetFieldValue(2, "^", "<" & Val(mcolParam("LESS_WEIGHT")), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    TempFromWeight = mcolParam("FROM_WEIGHT")
    TempToWeight = mcolParam("TO_WEIGHT")
   
    While TempFromWeight <= TempToWeight
        Call Cf.SetFieldValue(2, "^", TempFromWeight & "-" & (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1), ">")
        Call m_Headers1.AddField(Cf)
        TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
        ExportWidth = ExportWidth + 2
    Wend
    
   If Val(mcolParam("MORE_WEIGHT")) > 0 Then
        Call Cf.SetFieldValue(2, "^", ">" & Val(mcolParam("MORE_WEIGHT")), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN1")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME1"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN2")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME2"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN3")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME3"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN4")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME4"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    Call Cf.SetFieldValue(ExportWidth, "^", "จำนวน", ">")
    Call m_Headers0.AddField(Cf)
    '------------------------------------------------------------------------------------------------------------------------------
    
    'นน
    '------------------------------------------------------------------------------------------------------------------------------
    ExportWidth = 0
    If Val(mcolParam("LESS_WEIGHT")) > 0 Then
        Call Cf.SetFieldValue(2, "^", "<" & Val(mcolParam("LESS_WEIGHT")), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    TempFromWeight = mcolParam("FROM_WEIGHT")
    TempToWeight = mcolParam("TO_WEIGHT")
   
    While TempFromWeight <= TempToWeight
        Call Cf.SetFieldValue(2, "^", TempFromWeight & "-" & (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1), ">")
        Call m_Headers1.AddField(Cf)
        TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
        ExportWidth = ExportWidth + 2
    Wend
   If Val(mcolParam("MORE_WEIGHT")) > 0 Then
        Call Cf.SetFieldValue(2, "^", ">" & Val(mcolParam("MORE_WEIGHT")), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN1")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME1"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN2")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME2"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN3")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME3"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN4")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME4"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    Call Cf.SetFieldValue(ExportWidth, "^", "นน", ">")
    Call m_Headers0.AddField(Cf)
    '------------------------------------------------------------------------------------------------------------------------------
    
    'มูลค่า
    '------------------------------------------------------------------------------------------------------------------------------
    ExportWidth = 0
    If Val(mcolParam("LESS_WEIGHT")) > 0 Then
        Call Cf.SetFieldValue(2, "^", "<" & Val(mcolParam("LESS_WEIGHT")), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    TempFromWeight = mcolParam("FROM_WEIGHT")
    TempToWeight = mcolParam("TO_WEIGHT")
   
    While TempFromWeight <= TempToWeight
        Call Cf.SetFieldValue(2, "^", TempFromWeight & "-" & (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1), ">")
        Call m_Headers1.AddField(Cf)
        TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
        ExportWidth = ExportWidth + 2
    Wend
   If Val(mcolParam("MORE_WEIGHT")) > 0 Then
        Call Cf.SetFieldValue(2, "^", ">" & Val(mcolParam("MORE_WEIGHT")), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN1")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME1"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN2")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME2"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN3")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME3"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    If Val(mcolParam("PIG_STATUS_IN4")) > 0 Then
        Call Cf.SetFieldValue(2, "^", mcolParam("PIG_STATUS_IN_NAME4"), ">")
        Call m_Headers1.AddField(Cf)
        ExportWidth = ExportWidth + 2
    End If
    
    Call Cf.SetFieldValue(ExportWidth, "^", "มูลค่า", ">")
    Call m_Headers0.AddField(Cf)
    '------------------------------------------------------------------------------------------------------------------------------
    
   Set Cf = Nothing

  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
    
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
   Vsp.PaperSize = pprA4 'pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   If Val(mcolParam("PIG_STATUS1")) > 0 Then
        Call LoadWeightAmountPriceByCus(m_SumGroup1, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", BillingDocType2Set(mcolParam("BILL_SUBTYPE")), mcolParam("PIG_STATUS1"))
   End If
   If Val(mcolParam("PIG_STATUS2")) > 0 Then
        Call LoadWeightAmountPriceByCus(m_SumGroup2, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", BillingDocType2Set(mcolParam("BILL_SUBTYPE")), mcolParam("PIG_STATUS2"))
    End If
    If Val(mcolParam("PIG_STATUS3")) > 0 Then
        Call LoadWeightAmountPriceByCus(m_SumGroup3, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", BillingDocType2Set(mcolParam("BILL_SUBTYPE")), mcolParam("PIG_STATUS3"))
    End If
   
   Call GeneratePeriodData
   
   If Val(mcolParam("PIG_STATUS_IN1")) > 0 Then
        Call LoadSumWeightAmountPriceByCus(m_SumEx1, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", BillingDocType2Set(mcolParam("BILL_SUBTYPE")), mcolParam("PIG_STATUS_IN1"))
    End If
    If Val(mcolParam("PIG_STATUS_IN2")) > 0 Then
        Call LoadSumWeightAmountPriceByCus(m_SumEx2, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", BillingDocType2Set(mcolParam("BILL_SUBTYPE")), mcolParam("PIG_STATUS_IN2"))
   End If
   If Val(mcolParam("PIG_STATUS_IN3")) > 0 Then
        Call LoadSumWeightAmountPriceByCus(m_SumEx3, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", BillingDocType2Set(mcolParam("BILL_SUBTYPE")), mcolParam("PIG_STATUS_IN3"))
   End If
   If Val(mcolParam("PIG_STATUS_IN4")) > 0 Then
        Call LoadSumWeightAmountPriceByCus(m_SumEx4, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", BillingDocType2Set(mcolParam("BILL_SUBTYPE")), mcolParam("PIG_STATUS_IN4"))
    End If
   
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

Private Function GetKey(Di As CDoItem) As String
Dim OrderBy As Long
   
   If OrderBy = 1 Then
      GetKey = Di.CUSTOMER_ID & "-" & Di.PIG_STATUS
   ElseIf OrderBy = 2 Then
      GetKey = Di.CUSTOMER_ID & "-" & Di.PIG_STATUS
   Else
      GetKey = Di.CUSTOMER_ID & "-" & Di.PIG_STATUS
   End If
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Bd As CBillingDoc
Dim HaveValue As Boolean
Dim TempFromWeight  As Double
Dim TempToWeight As Double

Dim Di1 As CDoItem
Dim Di2 As CDoItem
Dim Di3 As CDoItem

Dim Di4 As CDoItem
Dim Di5 As CDoItem
Dim Di6 As CDoItem
Dim Di7 As CDoItem

Dim AssignKey As String

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j
    
   Call StartExportFile(Vsp)
   Vsp.StartDoc
    i = 0
    Set BodyCf = New CReportField
      
    Set Bd = New CBillingDoc
    Bd.BILLING_DOC_ID = -1
    Bd.FROM_DATE = mcolParam("FROM_DATE")
    Bd.TO_DATE = mcolParam("TO_DATE")
    Bd.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
    Call Bd.QueryData(5, Rs, iCount)
    i = 0
    
    While Not Rs.EOF
        i = i + 1
        Call Bd.PopulateFromRS(5, Rs)
         
        Call m_Details1.ClearField
        
        k = 0
'        k = k + 1
'        Set HeadCf = m_Headers1.Fields(k)
'        TempStr = I
'        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'        Call m_Details1.AddField(BodyCf)
         
         k = k + 1
        Set HeadCf = m_Headers1.Fields(k)
        TempStr = Bd.CUSTOMER_NAME '& "(" & BD.CUSTOMER_CODE & ")"
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         HaveValue = False
         
         Set Di4 = GetObject("CDoItem", m_SumEx1, Trim(Str(Bd.CUSTOMER_ID)))
         Set Di5 = GetObject("CDoItem", m_SumEx2, Trim(Str(Bd.CUSTOMER_ID)))
         Set Di6 = GetObject("CDoItem", m_SumEx3, Trim(Str(Bd.CUSTOMER_ID)))
         Set Di7 = GetObject("CDoItem", m_SumEx4, Trim(Str(Bd.CUSTOMER_ID)))
         'จำนวน
         '------------------------------------------------------------------------------------------------------------------------------
                 
         If Val(mcolParam("LESS_WEIGHT")) > 0 Then
             k = k + 1
             AssignKey = "LOW"
             Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.ITEM_AMOUNT + Di2.ITEM_AMOUNT + Di3.ITEM_AMOUNT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         TempFromWeight = mcolParam("FROM_WEIGHT")
         TempToWeight = mcolParam("TO_WEIGHT")
        
         While TempFromWeight <= TempToWeight
             'Call Cf.SetFieldValue(2, "^", TempFromWeight & "-" & (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1), ">")
             k = k + 1
            AssignKey = "IN-" & TempFromWeight
             Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.ITEM_AMOUNT + Di2.ITEM_AMOUNT + Di3.ITEM_AMOUNT
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
            Total1(k) = Total1(k) + Amt
         Wend
         
        If Val(mcolParam("MORE_WEIGHT")) > 0 Then
             k = k + 1
            AssignKey = "UP"
             Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.ITEM_AMOUNT + Di2.ITEM_AMOUNT + Di3.ITEM_AMOUNT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN1")) > 0 Then
             k = k + 1
            Amt = Di4.ITEM_AMOUNT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN2")) > 0 Then
             k = k + 1
            Amt = Di5.ITEM_AMOUNT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN3")) > 0 Then
             k = k + 1
            Amt = Di6.ITEM_AMOUNT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN4")) > 0 Then
             k = k + 1
            Amt = Di7.ITEM_AMOUNT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         '------------------------------------------------------------------------------------------------------------------------------
    
         'นน
         '------------------------------------------------------------------------------------------------------------------------------
         
         If Val(mcolParam("LESS_WEIGHT")) > 0 Then
             k = k + 1
            
            AssignKey = "LOW"
            Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.TOTAL_WEIGHT + Di2.TOTAL_WEIGHT + Di3.TOTAL_WEIGHT
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         TempFromWeight = mcolParam("FROM_WEIGHT")
         TempToWeight = mcolParam("TO_WEIGHT")
        
         While TempFromWeight <= TempToWeight
             'Call Cf.SetFieldValue(2, "^", TempFromWeight & "-" & (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1), ">")
             k = k + 1
            
            AssignKey = "IN-" & TempFromWeight
            Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.TOTAL_WEIGHT + Di2.TOTAL_WEIGHT + Di3.TOTAL_WEIGHT
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
            Total1(k) = Total1(k) + Amt
         Wend
         
        If Val(mcolParam("MORE_WEIGHT")) > 0 Then
             k = k + 1
            AssignKey = "UP"
            Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.TOTAL_WEIGHT + Di2.TOTAL_WEIGHT + Di3.TOTAL_WEIGHT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN1")) > 0 Then
             k = k + 1
            Amt = Di4.TOTAL_WEIGHT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN2")) > 0 Then
             k = k + 1
            Amt = Di5.TOTAL_WEIGHT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN3")) > 0 Then
             k = k + 1
            Amt = Di6.TOTAL_WEIGHT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN4")) > 0 Then
             k = k + 1
            Amt = Di7.TOTAL_WEIGHT
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         '------------------------------------------------------------------------------------------------------------------------------
         
                  'มูลค่า
         '------------------------------------------------------------------------------------------------------------------------------
         
         If Val(mcolParam("LESS_WEIGHT")) > 0 Then
             k = k + 1
            
            AssignKey = "LOW"
            Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.TOTAL_PRICE + Di2.TOTAL_PRICE + Di3.TOTAL_PRICE
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         TempFromWeight = mcolParam("FROM_WEIGHT")
         TempToWeight = mcolParam("TO_WEIGHT")
        
         While TempFromWeight <= TempToWeight
             'Call Cf.SetFieldValue(2, "^", TempFromWeight & "-" & (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1), ">")
             k = k + 1
            
            AssignKey = "IN-" & TempFromWeight
            Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.TOTAL_PRICE + Di2.TOTAL_PRICE + Di3.TOTAL_PRICE
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
            Total1(k) = Total1(k) + Amt
         Wend
         
        If Val(mcolParam("MORE_WEIGHT")) > 0 Then
             k = k + 1
            
            AssignKey = "UP"
            Set Di1 = GetObject("CDoItem", m_SumGroupNew1, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Set Di2 = GetObject("CDoItem", m_SumGroupNew2, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
             Set Di3 = GetObject("CDoItem", m_SumGroupNew3, Trim(Bd.CUSTOMER_ID & "-" & AssignKey))
            Amt = Di1.TOTAL_PRICE + Di2.TOTAL_PRICE + Di3.TOTAL_PRICE
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN1")) > 0 Then
             k = k + 1
            Amt = Di4.TOTAL_PRICE
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN2")) > 0 Then
             k = k + 1
            Amt = Di5.TOTAL_PRICE
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN3")) > 0 Then
             k = k + 1
            Amt = Di6.TOTAL_PRICE
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         
         If Val(mcolParam("PIG_STATUS_IN4")) > 0 Then
             k = k + 1
            Amt = Di7.TOTAL_PRICE
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            HaveValue = CheckHaveValue(HaveValue, Amt)
            Total1(k) = Total1(k) + Amt
         End If
         '------------------------------------------------------------------------------------------------------------------------------
        If HaveValue Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Call Vsp.AddTable(strFormat, "", TempStr2)
        End If
        
         Rs.MoveNext
         
      Wend

      Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Set BodyCf = Nothing

   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
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

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
      Else
         TempStr = FormatNumberToNull(Tot(j), 0)
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
Private Sub GeneratePeriodData()
Dim TempDi As CDoItem

Dim TempFromWeight As Double
Dim TempToWeight As Double
Dim AssignKey As String

Dim DiNew As CDoItem
    
    '111111111111111---------------------------------------------------------------------------------------------------------------------------------------------------
    If mcolParam("PIG_STATUS1") > 0 Then
        For Each TempDi In m_SumGroup1
            If TempDi.AVG_WEIGHT <= Val(mcolParam("LESS_WEIGHT")) Then
                AssignKey = "LOW"
            ElseIf TempDi.AVG_WEIGHT >= Val(mcolParam("MORE_WEIGHT")) Then
                AssignKey = "UP"
            Else
                TempFromWeight = mcolParam("FROM_WEIGHT")
                TempToWeight = mcolParam("TO_WEIGHT")
                
                While TempFromWeight < TempToWeight
                    If (TempDi.AVG_WEIGHT >= TempFromWeight) And (TempDi.AVG_WEIGHT <= (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1)) Then
                        AssignKey = "IN-" & TempFromWeight
                    End If
                    TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
                Wend
            End If
            
            Set DiNew = GetObject("CDoItem", m_SumGroupNew1, Trim(TempDi.CUSTOMER_ID & "-" & AssignKey), False)
            If DiNew Is Nothing Then
                Set DiNew = New CDoItem
                DiNew.CUSTOMER_ID = TempDi.CUSTOMER_ID
                DiNew.TOTAL_WEIGHT = TempDi.TOTAL_WEIGHT
                DiNew.ITEM_AMOUNT = TempDi.ITEM_AMOUNT
                DiNew.TOTAL_PRICE = TempDi.TOTAL_PRICE
                Call m_SumGroupNew1.Add(DiNew, Trim(TempDi.CUSTOMER_ID & "-" & AssignKey))
            Else
                DiNew.TOTAL_WEIGHT = DiNew.TOTAL_WEIGHT + TempDi.TOTAL_WEIGHT
                DiNew.ITEM_AMOUNT = DiNew.ITEM_AMOUNT + TempDi.ITEM_AMOUNT
                DiNew.TOTAL_PRICE = DiNew.TOTAL_PRICE + TempDi.TOTAL_PRICE
            End If
        Next TempDi
    End If
    '11111111---------------------------------------------------------------------------------------------------------------------------------------------------
    
  '2222222222---------------------------------------------------------------------------------------------------------------------------------------------------
    If mcolParam("PIG_STATUS2") > 0 Then
        For Each TempDi In m_SumGroup2
            If TempDi.AVG_WEIGHT <= Val(mcolParam("LESS_WEIGHT")) Then
                AssignKey = "LOW"
            ElseIf TempDi.AVG_WEIGHT >= Val(mcolParam("MORE_WEIGHT")) Then
                AssignKey = "UP"
            Else
                TempFromWeight = mcolParam("FROM_WEIGHT")
                TempToWeight = mcolParam("TO_WEIGHT")
                
                While TempFromWeight < TempToWeight
                    If (TempDi.AVG_WEIGHT >= TempFromWeight) And (TempDi.AVG_WEIGHT <= (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1)) Then
                        AssignKey = "IN-" & TempFromWeight
                     End If
                     TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
                Wend
            End If
            
            Set DiNew = GetObject("CDoItem", m_SumGroupNew2, Trim(TempDi.CUSTOMER_ID & "-" & AssignKey), False)
            If DiNew Is Nothing Then
                Set DiNew = New CDoItem
                DiNew.CUSTOMER_ID = TempDi.CUSTOMER_ID
                DiNew.TOTAL_WEIGHT = TempDi.TOTAL_WEIGHT
                DiNew.ITEM_AMOUNT = TempDi.ITEM_AMOUNT
                DiNew.TOTAL_PRICE = TempDi.TOTAL_PRICE
                Call m_SumGroupNew2.Add(DiNew, Trim(TempDi.CUSTOMER_ID & "-" & AssignKey))
            Else
                DiNew.TOTAL_WEIGHT = DiNew.TOTAL_WEIGHT + TempDi.TOTAL_WEIGHT
                DiNew.ITEM_AMOUNT = DiNew.ITEM_AMOUNT + TempDi.ITEM_AMOUNT
                DiNew.TOTAL_PRICE = DiNew.TOTAL_PRICE + TempDi.TOTAL_PRICE
            End If
        Next TempDi
    End If
    '222222222222---------------------------------------------------------------------------------------------------------------------------------------------------

'3333333333333---------------------------------------------------------------------------------------------------------------------------------------------------
    If mcolParam("PIG_STATUS3") > 0 Then
        For Each TempDi In m_SumGroup3
            If TempDi.AVG_WEIGHT <= Val(mcolParam("LESS_WEIGHT")) Then
                AssignKey = "LOW"
            ElseIf TempDi.AVG_WEIGHT >= Val(mcolParam("MORE_WEIGHT")) Then
                AssignKey = "UP"
            Else
                TempFromWeight = mcolParam("FROM_WEIGHT")
                TempToWeight = mcolParam("TO_WEIGHT")
                
                While TempFromWeight < TempToWeight
                    If (TempDi.AVG_WEIGHT >= TempFromWeight) And (TempDi.AVG_WEIGHT <= (TempFromWeight + Val(mcolParam("PERIOD_WEIGHT")) - 1)) Then
                        TempFromWeight = TempFromWeight + Val(mcolParam("PERIOD_WEIGHT"))
                    End If
                     AssignKey = "IN-" & TempFromWeight
                Wend
            End If
            
            Set DiNew = GetObject("CDoItem", m_SumGroupNew3, Trim(TempDi.CUSTOMER_ID & "-" & AssignKey), False)
            If DiNew Is Nothing Then
                Set DiNew = New CDoItem
                DiNew.CUSTOMER_ID = TempDi.CUSTOMER_ID
                DiNew.TOTAL_WEIGHT = TempDi.TOTAL_WEIGHT
                DiNew.ITEM_AMOUNT = TempDi.ITEM_AMOUNT
                DiNew.TOTAL_PRICE = TempDi.TOTAL_PRICE
                Call m_SumGroupNew1.Add(DiNew, Trim(TempDi.CUSTOMER_ID & "-" & AssignKey))
            Else
                DiNew.TOTAL_WEIGHT = DiNew.TOTAL_WEIGHT + TempDi.TOTAL_WEIGHT
                DiNew.ITEM_AMOUNT = DiNew.ITEM_AMOUNT + TempDi.ITEM_AMOUNT
                DiNew.TOTAL_PRICE = DiNew.TOTAL_PRICE + TempDi.TOTAL_PRICE
            End If
        Next TempDi
    End If
    '3333333333---------------------------------------------------------------------------------------------------------------------------------------------------

End Sub
