VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory007_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory007_2"
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
Private Const TITLE_HEAD_TABLE = "^3.5|^28.8;"
Private Const HEADER_FORMAT1 = "^1.5|^2.5|^7|^2.5|^2.5|^2|^2|^2|^1.8|^2.7|^3;"
Private Const HEADER_FORMAT2 = "^1.5|^2.5|^7|^2.5|^2.5|^2|^2|^2|^1.8|^2.7|^3|^4|^4;"
Private Const DETAIL_FORMAT1 = "^1.5|^2.5|<7|^2.5|^2.5|^2|>2|^2|^1.8|^2.7|<3;"
Private Const DETAIL_FORMAT2 = "^1.5|^2.5|<7|^2.5|^2.5|^2|>2|^2|^1.8|^2.7|<3|^4;"
Private Const FOOTER_FORMAT1 = "^1.5|^2.5|<7|^2.5|>2.5|>2|>2|>2|>1.8|>2.7|>3;"
Private Const FOOTER_FORMAT2 = "^1.5|^2.5|<7|^2.5|>2.5|>2|>2|>2|>1.8|>2.7|>3|>4;"
'ลำดับ|ประเภทวัตถุดิบ|วัตถุดิบ|ราคา/หน่วย|จำนวน|หน่วย|ราคารวม|สถานที่จัดเก็บ|เลขที่เอกสาร|วันที่เอกสาร|ประเภทเอกสาร|สถานะรายการ|ทะเบียนรถ;

Private Sub Class_Initialize()
   Set mcolParam = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
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
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
 
     strFormat = VSP_CalTable(HEADER_FORMAT2, mdblWid, alngX)
    strPrint = "ลำดับ|วันที่เอกสาร|วัตถุดิบ|ทะเบียน|ประเภท|ขนาดบรรจุ|จำนวน|หน่วย|รุ่นผลิต|วันเดือนปีผลิต|วันเดือนปีหมดอายุ|บริษัทจัดจำหน่าย|ทะเบียนรถ;"
  
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim FromDate As String
Dim ToDate As String

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

   FromDate = "จาก " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " ," & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & vbCrLf & FromDate & " - " & ToDate & ",  " & "กลุ่ม " & NVLS(Rs("PART_GROUP_NAME"), "______________________") & ", " & "สถานที่จัดเก็บ " & NVLS(Rs("LOCATION_NAME"), "______________________")
                                       
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
Dim RName As String
Dim cData As CImportItem
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim strFormat2 As String
Dim strPrint2 As String
Dim strFormat3 As String
Dim strPrint3 As String
Dim strFormat4 As String
Dim strPrint4 As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim Sum2 As Double
Dim LeftAmount As Double
Dim ValueAmount As Double
Dim PrevKey As String
Dim PrevKeyName As String
Dim NewKey As String
Dim NewKeyName As String
Dim CountSubtract As Long
   
   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
      
   Set cData = New CImportItem
     cData.IMPORT_ITEM_ID = -1
     cData.FROM_DATE = mcolParam("FROM_DATE")
     cData.TO_DATE = mcolParam("TO_DATE")
     cData.PART_NO = mcolParam("PART_NO")
    cData.PART_GROUP = mcolParam("PART_GROUP")
    cData.SUPPLIER_CODE = mcolParam("SUPPLIER_CODE")
    cData.LOCATION_ID = mcolParam("LOCATION_ID")
    If mcolParam("SUPPLIER_NAME_FLAG") = "Y" Then
      cData.OrderBy = 1
       cData.OrderType = 1
   ElseIf mcolParam("SUPPLIER_NAME_FLAG") = "N" Then
       cData.OrderBy = 2
       cData.OrderType = 1
   End If
      
   Call cData.QueryData(8888, Rs, iCount)

   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   Set cData = Nothing
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
'   ลำดับ|ประเภทวัตถุดิบ|วัตถุดิบ|ราคา/หน่วย|จำนวน|หน่วย|ราคารวม|สถานที่จัดเก็บ|เลขที่เอกสาร|วันที่เอกสาร|ประเภทเอกสาร|สถานะรายการ;
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Vsp.TableBorder = tbAll
   Sum1 = 0
   PrevKey = " "
   CountSubtract = 0
   If mcolParam("SUPPLIER_NAME_FLAG") = "Y" Then
      For I = 1 To iCount
          If PrevKey <> NVLS(Rs("SUP_NUM"), "") Then
           If I > 1 Then
               strFormat2 = VSP_CalTable(FOOTER_FORMAT2, mdblWid, alngX)
               strPrint2 = "รวม||||||" & FormatNumber(Sum1) & "||||||;"
               Call Vsp.AddTable(strFormat2, "", strPrint2)
         End If
        
             strFormat3 = VSP_CalTable(TITLE_HEAD_TABLE, mdblWid, alngX)
             strPrint3 = NVLS(Rs("SUP_NUM"), "") & "|"
             strPrint3 = strPrint3 & NVLS(Rs("SUP_NAME"), "") & "|"
            Call Vsp.AddTable(strFormat3, "", strPrint3)
         
          Sum1 = 0
         End If
   
          strFormat = VSP_CalTable(DETAIL_FORMAT2, mdblWid, alngX)
          
          strPrint = (I - CountSubtract) & "|"
         strPrint = strPrint & InternalDateToStringEx4(NVLS(Rs("DOCUMENT_DATE"), "")) & "|"
         strPrint = strPrint & NVLS(Rs("PART_DESC"), "") & " (" & NVLS(Rs("PART_NO"), "") & ")|"
        strPrint = strPrint & NVLS(Rs("REGIST_NUM"), "") & "|"
         strPrint = strPrint & NVLS(Rs("PART_TYPE_NAME"), "") & " (" & NVLS(Rs("PART_TYPE_NO"), "") & ")|"

         strPrint = strPrint & NVLS(Rs("DRUG_SIZE"), "") & "|"
         strPrint = strPrint & FormatNumber(NVLD(Rs("IMPORT_AMOUNT"), 0), "-") & "|"
         strPrint = strPrint & NVLS(Rs("UNIT_NAME"), "") & "|"
         strPrint = strPrint & "|"
         strPrint = strPrint & "|"
        strPrint = strPrint & "|"
'm_PartItem.DRUG_SIZE
         strPrint = strPrint & NVLS(Rs("SUP_NAME"), "") & "|"

       
        
        
         Sum1 = Sum1 + NVLD(Rs("IMPORT_AMOUNT"), 0)
       
         If Val(NVLD(Rs("IMPORT_AMOUNT"), 0)) > 0 Then
            CountSubtract = CountSubtract + 1
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         PrevKey = NVLS(Rs("SUP_NUM"), "")
         PrevKeyName = NVLS(Rs("SUP_NAME"), "")
         Rs.MoveNext
         If Rs.EOF Then
            Call AdjustPage(Vsp, strFormat, strPrint, 1500)
         End If

      Next I
     End If
     
     
     CountSubtract = 0
     If mcolParam("SUPPLIER_NAME_FLAG") = "N" Then
      For I = 1 To iCount
         strFormat = VSP_CalTable(DETAIL_FORMAT2, mdblWid, alngX)
          strPrint = (I - CountSubtract) & "|"
         strPrint = strPrint & InternalDateToStringEx4(NVLS(Rs("DOCUMENT_DATE"), "")) & "|"
         strPrint = strPrint & NVLS(Rs("PART_DESC"), "") & " (" & NVLS(Rs("PART_NO"), "") & ")|"
        strPrint = strPrint & NVLS(Rs("REGIST_NUM"), "") & "|"
'         strPrint = strPrint & NVLS(Rs("PART_TYPE_NAME"), "") & " (" & NVLS(Rs("PART_TYPE_NO"), "") & ")|"
        strPrint = strPrint & NVLS(Rs("PART_TYPE_NAME"), "") & "|"
         strPrint = strPrint & NVLS(Rs("DRUG_SIZE"), "") & "|"
         strPrint = strPrint & FormatNumber(NVLD(Rs("IMPORT_AMOUNT"), 0), "-") & "|"
         strPrint = strPrint & NVLS(Rs("UNIT_NAME"), "") & "|"
         strPrint = strPrint & "|"
         strPrint = strPrint & "|"
        strPrint = strPrint & "|"

         strPrint = strPrint & NVLS(Rs("SUP_NAME"), "") & "|"

         Sum1 = Sum1 + NVLD(Rs("IMPORT_AMOUNT"), 0)
         
         If Val(NVLD(Rs("IMPORT_AMOUNT"), 0)) > 0 Then
            CountSubtract = CountSubtract + 1
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         Rs.MoveNext
         If Rs.EOF Then
            Call AdjustPage(Vsp, strFormat, strPrint, 1500)
         End If
 
      Next I
     End If
          strFormat4 = VSP_CalTable(FOOTER_FORMAT2, mdblWid, alngX)
          strPrint4 = "รวม||||||" & FormatNumber(Sum1) & "||||||;"
         Call Vsp.AddTable(strFormat4, "", strPrint4)
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
      Set Rs = Nothing
   End If
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function


Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
' Vsp = vbCrLf & "                       ผู้จัดทำ _________________________________                                           ผู้จัดการฟาร์ม  _________________________                                      สัตวแพทย์ควบคุมฟาร์ม  ______________________" & _
'            vbCrLf & "                         (                                                                       )                                            (                                                                          )                                        (                                                                          )" & _
'            vbCrLf & "                                      ว/ด/ป   ___ / ___ / ___                                                                                  ว/ด/ป   ___ / ___ / ___                                                                                         ว/ด/ป   ___ / ___ / ___ "
' If Not mblnEndPage Then Exit Sub
End Sub

Private Sub VSP_EndPage()
   
   Vsp = vbCrLf & "                       ผู้จัดทำ _________________________________                                           ผู้จัดการฟาร์ม  _________________________                                      สัตวแพทย์ควบคุมฟาร์ม  ______________________" & _
            vbCrLf & "                         (                                                                       )                                            (                                                                          )                                        (                                                                          )" & _
            vbCrLf & "                                      ว/ด/ป   ___ / ___ / ___                                                                                  ว/ด/ป   ___ / ___ / ___                                                                                         ว/ด/ป   ___ / ___ / ___ "
 If Not mblnEndPage Then Exit Sub
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


