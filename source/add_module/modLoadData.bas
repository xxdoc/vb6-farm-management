Attribute VB_Name = "modLoadData"
Option Explicit
' Test test test
Public Sub GetMonthAllGoods(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional ShowOutlay As Long)
On Error GoTo ErrorHandler
Dim Di As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
   
   Set Di = New CDoItem
   Set Rs = New ADODB.Recordset
   
    Di.DO_ITEM_ID = -1
   Di.FROM_DATE = FromDate
   Di.TO_DATE = ToDate
   Call Di.QueryData(40, Rs, itemcount)
    
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(40, Rs)
      
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData)
'      End If
'
    If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE & "-" & TempData.YYYYMM))
         Call Cl.Add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.YYYYMM))
  '       ''debug.print (Trim(TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE & "-" & TempData.YYYYMM))
      End If

      
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
  
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub


Public Sub GetDistinctProductStatus(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional ShowOutlay As Long)
On Error GoTo ErrorHandler
Dim Di As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
   
   Set Di = New CDoItem
   Set Rs = New ADODB.Recordset
   
    Di.DO_ITEM_ID = -1
   Di.FROM_DATE = FromDate
   Di.TO_DATE = ToDate
   Call Di.QueryData(41, Rs, itemcount)
    
   
   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
   End If
   
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(41, Rs)
      
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData)
'      End If
'
    If Not (Cl Is Nothing) Then
  If TempData.PRODUCT_STATUS_NO = "" Then
         Call Cl.Add(TempData, Trim(TempData.PART_TYPE_NO & "-" & TempData.PART_TYPE_NAME))
    Else
      Call Cl.Add(TempData, Trim(TempData.PRODUCT_STATUS_NO & "-" & TempData.PRODUCT_STATUS_NAME))
    End If
      End If

      
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
  
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub
Public Sub GetProductStatusTypeYYYYMM(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim Di As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
   
   Set Di = New CDoItem
   Set Rs = New ADODB.Recordset
   
    Di.DO_ITEM_ID = -1
   Di.FROM_DATE = FromDate
   Di.TO_DATE = ToDate
   Call Di.QueryData(43, Rs, itemcount)
    
   
   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
   End If
   
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(43, Rs)
      
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData)
'      End If
'
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.REVENUE_NO & "-" & TempData.PRODUCT_STATUS_NO & "-" & TempData.PART_TYPE_NO & "-" & TempData.PIG_FLAG & "-" & TempData.YYYYMM))
         ''debug.print Trim(TempData.PRODUCT_STATUS_NO & "-" & TempData.PART_TYPE_NO & "-" & TempData.PIG_FLAG & "-" & TempData.YYYYMM)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
  
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub

Public Sub GetAllProduct(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional ShowOutlay As Long)
On Error GoTo ErrorHandler
Dim Di As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
   
   Set Di = New CDoItem
   Set Rs = New ADODB.Recordset
   
    Di.DO_ITEM_ID = -1
   Di.FROM_DATE = FromDate
   Di.TO_DATE = ToDate
   Call Di.QueryData(39, Rs, itemcount)
    
   
   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
   End If
   
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(39, Rs)
      
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData)
'      End If
'
    If Not (Cl Is Nothing) Then
  
         Call Cl.Add(TempData, Trim(TempData.DO_ITEM_ID & "-" & TempData.YYYYMM))
   
      End If

      
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
  
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub
Public Sub InitThaiMonth(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("มกราคม"))
   C.ItemData(1) = 1

   C.AddItem (MapText("กุมภาพันธ์"))
   C.ItemData(2) = 2

C.AddItem (MapText("มีนาคม"))
   C.ItemData(3) = 3

C.AddItem (MapText("เมษายน"))
   C.ItemData(4) = 4

C.AddItem (MapText("พฤษภาคม"))
   C.ItemData(5) = 5

C.AddItem (MapText("มิถุนายน"))
   C.ItemData(6) = 6

C.AddItem (MapText("กรกฎาคม"))
   C.ItemData(7) = 7

C.AddItem (MapText("สิงหาคม"))
   C.ItemData(8) = 8

C.AddItem (MapText("กันยายน"))
   C.ItemData(9) = 9

C.AddItem (MapText("ตุลาคม"))
   C.ItemData(10) = 10

C.AddItem (MapText("พฤศจิกายน"))
   C.ItemData(11) = 11

   C.AddItem (MapText(" ธันวาคม"))
   C.ItemData(12) = 12
End Sub
Public Sub InitFamilyStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("มีครอบครัว")
   C.ItemData(1) = 1
   
   C.AddItem ("โสด")
   C.ItemData(2) = 2
   
   C.AddItem ("หม้าย")
   C.ItemData(3) = 3
End Sub
Public Sub LoadDrug(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDrug
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDrug
Dim I As Long

   Set D = New CDrug
   Set Rs = New ADODB.Recordset
   
   D.DRUG_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDrug
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DRUG_NAME)
         C.ItemData(I) = TempData.DRUG_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.DRUG_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBloodSpec(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CBloodSpec
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBloodSpec
Dim I As Long

   Set D = New CBloodSpec
   Set Rs = New ADODB.Recordset
   
   D.BLOOD_SPEC_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBloodSpec
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SPEC_NAME)
         C.ItemData(I) = TempData.BLOOD_SPEC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.BLOOD_SPEC_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPeriodDesc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDSheetItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDSheetItem
Dim I As Long

   Set D = New CDSheetItem
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData2(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDSheetItem
      Call TempData.PopulateFromRS2(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PERIOD_DESC)
         C.ItemData(I) = TempData.DSHEET_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadXCollection(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CXCollection
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CXCollection
Dim I As Long

   Set D = New CXCollection
   Set Rs = New ADODB.Recordset
   
   D.X_COLLECTION_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New CXCollection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CXCollection
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.X_COLLECTION_NAME)
         C.ItemData(I) = TempData.X_COLLECTION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.X_COLLECTION_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Public Sub LoadYCollection(C As ComboBox, Optional Cl As Collection = Nothing)
'On Error GoTo ErrorHandler
'Dim ItemCount As Long
'Dim Rs As ADODB.Recordset
'Dim TempData As CYCollection
'Dim i As Long
'
'   Set D = New CYCollection
'   Set Rs = New ADODB.Recordset
'
'   D.Y_COLLECTION_ID = -1
'   Call D.QueryData(Rs, ItemCount)
'
'   If Not (C Is Nothing) Then
'      C.Clear
'      i = 0
'      C.AddItem ("")
'   End If
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New CYCollection
'   End If
'   While Not Rs.EOF
'      i = i + 1
'      Set TempData = New CYCollection
'      Call TempData.PopulateFromRS(Rs)
'
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.Y_COLLECTION_NAME)
'         C.ItemData(i) = TempData.Y_COLLECTION_ID
'      End If
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.Add(TempData, Str(TempData.Y_COLLECTION_ID))
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'
'   Set Rs = Nothing
'   Set D = Nothing
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub

Public Sub InitUserGroupOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(1) = 1
End Sub

Public Sub InitUserStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใช้งานได้")
   C.ItemData(1) = 1

   C.AddItem ("ถูกระงับ")
   C.ItemData(2) = 2
End Sub

Public Sub InitLoginOrderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่ล็อคอิน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อผู้ใช้"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport4_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("หมายเลขวัตถุดิบ"))
   C.ItemData(2) = 2

   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(3) = 3

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(4) = 4
End Sub

Public Sub InitReport4_5Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport4_14Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสโรงเรือน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อโรงเรือน"))
   C.ItemData(2) = 2

   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(3) = 3

   C.AddItem (MapText("ชื่อวัตถุดิบ"))
   C.ItemData(4) = 4
End Sub

Public Sub InitReport4_7Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
   C.ItemData(1) = 5

   C.AddItem (MapText("หมายเลขวัตถุดิบ"))
   C.ItemData(2) = 6

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(3) = 7

   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(4) = 8

   C.AddItem (MapText("รหัสสถานที่จัดเก็บ"))
   C.ItemData(5) = 9
End Sub

Public Sub InitReport4_15Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
'   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
'   C.ItemData(1) = 1
'
'   C.AddItem (MapText("หมายเลขวัตถุดิบ"))
'   C.ItemData(2) = 2
'
'   C.AddItem (MapText("เลขที่เอกสาร"))
'   C.ItemData(3) = 3

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสโรงเรือน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อโรงเรือน"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport5_2Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_5Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_6Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสโรงเรือน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อโรงเรือน"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport5_14Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("สัปดาห์เกิด"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_17Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("ประเภทสุกร"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_18Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport6_3_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 2

   C.AddItem (MapText("สัปดาห์เกิด"))
   C.ItemData(1) = 5
End Sub

Public Sub InitReport6_3_2Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 6

   C.AddItem (MapText("อายุ"))
   C.ItemData(1) = 6
End Sub

Public Sub InitReport6_3_3Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport6_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 2

   C.AddItem (MapText("รหัสสถานะสุกร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อสถานะสุกร"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport6_3_11Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = -1

   C.AddItem (MapText("สัปดาห์เกิด"))
   C.ItemData(1) = 1

   C.AddItem (MapText("สถานะสุกร"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport6_3_14Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = -1

   C.AddItem (MapText("สัปดาห์เกิด"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport6_3_13Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = -1

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 8

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 9
End Sub

Public Sub InitReport6_3_12Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = -1

   C.AddItem (MapText("ช่วงอายุสุกร"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport6_3_7Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = -1

   C.AddItem (MapText("ประเภทลูกค้า"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport6_3_6Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 2

   C.AddItem (MapText("รหัสรายรับ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อรายรับ"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport6_2Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 4

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(1) = 3
End Sub

Public Sub InitReport6_5Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("สัปดาห์เกิดสุกร"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport6_14Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสโรงเรือน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อโรงเรือน"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport6_22Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("สัปดาห์เกิด"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport6_23Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 4

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 4

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 5
End Sub

Public Sub InitReport6_26Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("วันที่นำฝาก"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เลขที่บัญชี"))
   C.ItemData(2) = 2
End Sub
Public Sub InitReport6_27Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport6_30Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("ประเภทการชำระเงิน"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport6_24Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 2

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(1) = 2

   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(2) = 1
End Sub

Public Sub InitReport8_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
End Sub

Public Sub InitReport6_7Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

'   C.AddItem (MapText("สัปดาห์เกิดสุกร"))
'   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อโรงเรือน"))
   C.ItemData(1) = 2

   C.AddItem (MapText("รหัสโรงเรือน"))
   C.ItemData(2) = 3
End Sub

Public Sub InitReport5_15Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_11Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสโรงเรือน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อโรงเรือน"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport5_12Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("สัปดาห์เกิด"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_4Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อประเภทวัตถุดิบ"))
   C.ItemData(2) = 2
End Sub

Public Sub InitUserOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อผู้ใช้")
   C.ItemData(1) = 1

   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitDocumentType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใบนำเข้า")
   C.ItemData(1) = 1

   C.AddItem ("ใบเบิกวัตถุดิบ")
   C.ItemData(2) = 2

   C.AddItem ("ใบโอนวัตถุดิบ")
   C.ItemData(3) = 3

   C.AddItem ("ใบปรับยอด")
   C.ItemData(4) = 4
End Sub

Public Sub InitCommitStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("คำนวณแล้ว")
   C.ItemData(1) = 1

   C.AddItem ("ยังไม่คำนวณ")
   C.ItemData(2) = 2
End Sub

Public Sub InitBillSubType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ขายเชื่อ")
   C.ItemData(1) = 10

   C.AddItem ("ขายสด")
   C.ItemData(2) = 13
End Sub

Public Sub InitBillingBillSubType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ขายเชื่อ")
   C.ItemData(1) = 1

   C.AddItem ("ขายสด")
   C.ItemData(2) = 2
End Sub

Public Sub InitCustomerOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitParameterOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่พารามิเตอร์"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่พารามิเตอร์"))
   C.ItemData(2) = 2
End Sub

Public Sub InitBatchOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่แบต"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่แบต"))
   C.ItemData(2) = 2
End Sub

Public Sub InitSupplierOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสซัพพลายเออร์"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อซัพพลายเออร์"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitEmployeeOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อ"))
   C.ItemData(2) = 2

   C.AddItem (MapText("นามสกุล"))
   C.ItemData(3) = 3

   C.AddItem (MapText("ตำแหน่ง"))
   C.ItemData(4) = 4
End Sub
'===

Public Sub InitPartItemOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขวัตถุดิบ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อวัตถุดิบ"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitInventoryDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่บิลรับของ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่รับวัตถุดิบ"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitBillingDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(3) = 3
End Sub
'===
Public Sub InitBillingDocOtherFilterOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("PO ที่ยังไม่อนุมัติ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("PO ที่อนุมัติแล้ว"))
   C.ItemData(2) = 2

   C.AddItem (MapText("เอกสารที่สร้างโดยไม่มี PO และยังไม่อนุมัติ"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("เอกสารที่สร้างโดยไม่มี PO และอนุมัติแล้ว"))
   C.ItemData(4) = 4
End Sub

Public Sub InitPaymentOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่บัญชี"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitBillingDocExOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2

   C.AddItem (MapText("รหัสซับพลายเออร์"))
   C.ItemData(3) = 3
End Sub
'===

Public Sub InitBillingDocCapitalOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2

End Sub
'===

Public Sub InitPigDoc1OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ใบเกิด"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เกิด"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitInventoryDoc2OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขใบเบิก"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เบิก"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitInventoryDoc3OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขใบโอน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่โอน"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitPigDoc4OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขใบนำเข้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่ใบนำเข้า"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitPigDoc3OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขใบปรับยอด"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่ปรับยอด"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitPigWeekOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ปี"))
   C.ItemData(1) = 1
End Sub
'===

Public Sub LoadUserGroup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserGroup
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserGroup
Dim I As Long

   Set D = New CUserGroup
   Set Rs = New ADODB.Recordset
   
   D.GROUP_ID = -1
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GROUP_NAME)
         C.ItemData(I) = TempData.GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.GROUP_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadUserAccount(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserAccount
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserAccount
Dim I As Long

   Set D = New CUserAccount
   Set Rs = New ADODB.Recordset
   
   D.GROUP_ID = -1
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.USER_NAME)
         C.ItemData(I) = TempData.USER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.USER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadCountry(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCountry
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCountry
Dim I As Long

   Set D = New CCountry
   Set Rs = New ADODB.Recordset
   
   D.COUNTRY_ID = -1
   D.CONTINENT_ID = -1
   D.COUNTRY_NAME = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCountry
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.COUNTRY_NAME)
         C.ItemData(I) = TempData.COUNTRY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.COUNTRY_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadLocation(C As ComboBox, Optional Cl As Collection = Nothing, Optional Area As Long = -1, Optional SaleFlag As String = "N", Optional LocationID As Long = -1, Optional OrderBy As Long = 2)
On Error GoTo ErrorHandler
Static D As CLocation
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CLocation
Dim I As Long
   
   If Rs Is Nothing Then
      Set D = New CLocation
      Set Rs = New ADODB.Recordset
   
      D.LOCATION_ID = LocationID
      D.LOCATION_TYPE = Area
      
      D.SALE_FLAG = SaleFlag
      D.OrderBy = OrderBy
      Call D.QueryData(Rs, itemcount)
   Else
      If (D.LOCATION_ID <> LocationID) Or _
         (D.LOCATION_TYPE <> Area) Or _
         (D.SALE_FLAG <> SaleFlag) Then
      
         D.LOCATION_ID = LocationID
         D.LOCATION_TYPE = Area
         D.SALE_FLAG = SaleFlag
         D.OrderBy = OrderBy
         Call D.QueryData(Rs, itemcount)
      End If
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLocation
      Call TempData.PopulateFromRS(1, Rs)
      
'      If TempData.LOCATION_ID = 492 Then
'         ''debug.print
'      End If
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.LOCATION_NAME)
         C.ItemData(I) = TempData.LOCATION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)))
'''debug.print TempData.LOCATION_ID
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Not Rs.BOF Then
      Rs.MoveFirst
   End If
   
   'Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadLocationByCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional Area As Long = -1, Optional SaleFlag As String = "N", Optional LocationID As Long = -1, Optional OrderBy As Long = 2, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CLocation
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLocation
Dim I As Long

   Set D = New CLocation
   Set Rs = New ADODB.Recordset
   
   D.LOCATION_ID = LocationID
   D.LOCATION_TYPE = Area
   D.SALE_FLAG = SaleFlag
   D.OrderBy = OrderBy
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLocation
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.LOCATION_NAME)
         C.ItemData(I) = TempData.LOCATION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.Add(TempData, Trim(TempData.LOCATION_NO))
         ElseIf KeyType = 2 Then
            Call Cl.Add(TempData, Trim(TempData.LOCATION_NO & "-" & TempData.LOCATION_TYPE))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRevenueType(C As ComboBox, Optional Cl As Collection = Nothing, Optional RevenueID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CRevenueType
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRevenueType
Dim I As Long

   Set D = New CRevenueType
   Set Rs = New ADODB.Recordset
   
   D.REVENUE_TYPE_ID = RevenueID
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CRevenueType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.REVENUE_NAME)
         C.ItemData(I) = TempData.REVENUE_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.REVENUE_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadRevenueTypeCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional RevenueID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CRevenueType
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRevenueType
Dim I As Long

   Set D = New CRevenueType
   Set Rs = New ADODB.Recordset
   
   D.REVENUE_TYPE_ID = RevenueID
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CRevenueType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.REVENUE_NAME)
         C.ItemData(I) = TempData.REVENUE_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.REVENUE_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadLocationEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional HouseGroupID As Long)
On Error GoTo ErrorHandler
Dim D As CHouseGroup
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLocation
Dim I As Long
Dim IsOK As Boolean
Dim iCount As Long
Dim Hgi As CHGroupItem

   Set D = New CHouseGroup
   Set Rs = New ADODB.Recordset
   
   D.HOUSE_GROUP_ID = HouseGroupID
   D.EXTRA_FLAG = ""
   D.QueryFlag = 1
   Call glbMaster.QueryHouseGroup(D, Rs, iCount, IsOK, glbErrorLog)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   For Each Hgi In D.HGroupItems
      I = I + 1
      If Hgi.SELECT_FLAG = "Y" Then
         Set TempData = New CLocation
         TempData.LOCATION_ID = Hgi.LOCATION_ID
         TempData.LOCATION_NO = Hgi.LOCATION_NO
         TempData.LOCATION_NAME = Hgi.LOCATION_NAME
      
         If Not (C Is Nothing) Then
            C.AddItem (TempData.LOCATION_NAME)
            C.ItemData(I) = TempData.LOCATION_ID
         End If
         
         If Not (Cl Is Nothing) Then
            Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)))
         End If
         
         Set TempData = Nothing
      End If
   Next Hgi
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadCustomerGrade(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCustomerGrade
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomerGrade
Dim I As Long

   Set D = New CCustomerGrade
   Set Rs = New ADODB.Recordset
   
   D.CSTGRADE_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomerGrade
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CSTGRADE_NAME)
         C.ItemData(I) = TempData.CSTGRADE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.CSTGRADE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadCustomerType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCustomerType
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomerType
Dim I As Long

   Set D = New CCustomerType
   Set Rs = New ADODB.Recordset
   
   D.CSTTYPE_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomerType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CSTTYPE_NAME)
         C.ItemData(I) = TempData.CSTTYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.CSTTYPE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadSupplierGrade(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSupplierGrade
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierGrade
Dim I As Long

   Set D = New CSupplierGrade
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_GRADE_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierGrade
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_GRADE_NAME)
         C.ItemData(I) = TempData.SUPPLIER_GRADE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.SUPPLIER_GRADE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadSupplierType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSupplierType
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierType
Dim I As Long

   Set D = New CSupplierType
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_TYPE_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_TYPE_NAME)
         C.ItemData(I) = TempData.SUPPLIER_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.SUPPLIER_TYPE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadSupplierStatus(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSupplierStatus
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierStatus
Dim I As Long

   Set D = New CSupplierStatus
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_STATUS_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierStatus
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_STATUS_NAME)
         C.ItemData(I) = TempData.SUPPLIER_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.SUPPLIER_STATUS_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadPosition(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CEmpPosition
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEmpPosition
Dim I As Long

   Set D = New CEmpPosition
   Set Rs = New ADODB.Recordset
   
   D.POSITION_ID = -1
   D.POSITION_NAME = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEmpPosition
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.POSITION_DESC)
         C.ItemData(I) = TempData.POSITION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadYearSeq(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CYearSeq
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CYearSeq
Dim I As Long

   Set D = New CYearSeq
   Set Rs = New ADODB.Recordset
   
   D.YEAR_SEQ_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CYearSeq
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.YEAR_NO)
         C.ItemData(I) = TempData.YEAR_SEQ_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigStatus(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CProductStatus
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CProductStatus
Dim I As Long

   Set D = New CProductStatus
   Set Rs = New ADODB.Recordset
   
   D.PRODUCT_STATUS_ID = -1
   D.PRODUCT_STATUS_NAME = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CProductStatus
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PRODUCT_STATUS_NAME)
         C.ItemData(I) = TempData.PRODUCT_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadStatusType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CStatusType
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStatusType
Dim I As Long

   Set D = New CStatusType
   Set Rs = New ADODB.Recordset
   
   D.STATUS_TYPE_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStatusType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STATUS_TYPE_NAME)
         C.ItemData(I) = TempData.STATUS_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartType(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartGroupID As Long, Optional PartTypeID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPartType
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartType
Dim I As Long

   Set D = New CPartType
   Set Rs = New ADODB.Recordset
   
   D.PART_TYPE_ID = PartTypeID
   D.PART_TYPE_NAME = ""
   D.PART_GROUP_ID = PartGroupID
   D.PART_NO = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPartType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_TYPE_NAME)
         C.ItemData(I) = TempData.PART_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadExposeType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim D As CHGroupItem
Dim TempData As CHGroupItem
Dim I As Long

   Set Rs = New ADODB.Recordset
   Set D = New CHGroupItem
'   D.EXPOSE_TYPE_ID = ExposeTypeID
'   D.EXPOSE_TYPE_NAME = ""
   Call D.QueryData(3, Rs, itemcount)
   
'   If Not (C Is Nothing) Then
'      C.Clear
'      I = 0
'      C.AddItem ("")
'   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CHGroupItem
      Call TempData.PopulateFromRS(3, Rs)
   
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.EXPOSE_TYPE_NAME)
'         C.ItemData(I) = TempData.EXPOSE_TYPE_ID
'      End If
      
      If Not (Cl Is Nothing) Then
         TempData.EXIST_FLAG = "N"
         Call Cl.Add(TempData, Trim(Str(TempData.HOUSE_GROUP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigStatusInGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional GroupID As Long)
On Error GoTo ErrorHandler
Dim D As CSGroupItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSGroupItem
Dim I As Long

   Set D = New CSGroupItem
   Set Rs = New ADODB.Recordset
   
   D.SGROUP_ITEM_ID = -1
   D.STATUS_GROUP_ID = GroupID
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSGroupItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STATUS_NAME)
         C.ItemData(I) = TempData.STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.STATUS_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartGroup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPartGroup
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartGroup
Dim I As Long

   Set D = New CPartGroup
   Set Rs = New ADODB.Recordset
   
   D.PART_GROUP_ID = -1
   D.PART_GROUP_NAME = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPartGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_GROUP_NAME)
         C.ItemData(I) = TempData.PART_GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_GROUP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadHouseGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExtraFlag = "")
On Error GoTo ErrorHandler
Dim D As CHouseGroup
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CHouseGroup
Dim I As Long

   Set D = New CHouseGroup
   Set Rs = New ADODB.Recordset
   
   D.HOUSE_GROUP_ID = -1
   D.EXTRA_FLAG = ExtraFlag
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CHouseGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.HOUSE_GROUP_NAME)
         C.ItemData(I) = TempData.HOUSE_GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.HOUSE_GROUP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadHGroupItem(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CHGroupItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CHGroupItem
Dim I As Long
Dim test As String

   Set D = New CHGroupItem
   Set Rs = New ADODB.Recordset
   
   D.HOUSE_GROUP_ID = -1
 '  D.EXTRA_FLAG = ExtraFlag
 '  D.OrderBy = 1
 '  D.OrderType = 1
   Call D.QueryData(2, Rs, itemcount)
   
'   If Not (C Is Nothing) Then
'      C.Clear
'      I = 0
'      C.AddItem ("")
'   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CHGroupItem
      Call TempData.PopulateFromRS(2, Rs)
   
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.HOUSE_GROUP_NAME)
'         C.ItemData(I) = TempData.HOUSE_GROUP_ID
'      End If
      'เปลี่ยนเอา comment ออก version 295     เนื่องจากติด error ตรง add collection ตอน call  LoadHGroupItem ใน initDoc CReportInventoryDoc0069 หากอันอื่นเกิดปัญหา ให้สลับใช้ comment อันข้างล่าง
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)))
'        test = Str(TempData.LOCATION_ID)

'         ''debug.print (TempData & "-" & Trim(Str(TempData.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
'   ''debug.print test
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadStatusGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExtraFlag = "")
On Error GoTo ErrorHandler
Dim D As CStatusGroup
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStatusGroup
Dim I As Long

   Set D = New CStatusGroup
   Set Rs = New ADODB.Recordset
   
   D.STATUS_GROUP_ID = -1
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStatusGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STATUS_GROUP_NAME)
         C.ItemData(I) = TempData.STATUS_GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.STATUS_GROUP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadYearWeek(C As ComboBox, Optional Cl As Collection = Nothing, Optional YearSeqID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CYearWeek
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CYearWeek
Dim I As Long

   Set D = New CYearWeek
   Set Rs = New ADODB.Recordset
   
   D.YEAR_WEEK_ID = -1
   D.YEAR_SEQ_ID = YearSeqID
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CYearWeek
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.WEEK_NO)
         C.ItemData(I) = TempData.YEAR_WEEK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.YEAR_WEEK_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadYearWeekEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional YearSeqID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CYearWeek
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CYearWeek
Dim I As Long

   Set D = New CYearWeek
   Set Rs = New ADODB.Recordset
   
   D.YEAR_WEEK_ID = -1
   D.YEAR_SEQ_ID = YearSeqID
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CYearWeek
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.WEEK_NO)
         C.ItemData(I) = TempData.YEAR_WEEK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YEAR_NO & "-" & TempData.WEEK_NO)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctYearWeek(C As ComboBox, Optional Cl As Collection = Nothing, Optional YearSeqID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CYearWeek
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CYearWeek
Dim I As Long

   Set D = New CYearWeek
   Set Rs = New ADODB.Recordset
   
   D.YEAR_WEEK_ID = -1
   D.YEAR_SEQ_ID = YearSeqID
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CYearWeek
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.WEEK_NO)
         C.ItemData(I) = TempData.YEAR_WEEK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.WEEK_NO)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Function CompareInt(Key1 As Long, Key2 As Long) As Boolean
   If Key2 <= 0 Then
      CompareInt = True
   Else
      CompareInt = (Key1 = Key2)
   End If
End Function

Public Function CompareStr(Key1 As String, Key2 As String) As Boolean
   If Len(Key2) <= 0 Then
      CompareStr = True
   Else
      CompareStr = (Key1 = Key2)
   End If
End Function
Public Sub LoadPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartTypeID As Long = -1, Optional PigFlag As String = "N", Optional PigType As String = "", Optional SpecificFlag As String = "", Optional ParentID As Long = -1, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Static D As CPartItem
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CPartItem
Dim I As Long
   
      
   If Rs Is Nothing Then
      Set D = New CPartItem
      Set Rs = New ADODB.Recordset
   
      D.PART_ITEM_ID = -1
      D.PART_TYPE = PartTypeID
      D.PIG_FLAG = PigFlag
      D.PIG_TYPE = PigType
      D.SPECIFIC_FLAG = SpecificFlag
      D.PARENT_ID = ParentID
      Call D.QueryData(1, Rs, itemcount)
   Else
      If (D.PART_TYPE <> PartTypeID) Or _
         (D.PIG_FLAG <> PigFlag) Or _
         (D.SPECIFIC_FLAG <> SpecificFlag) Or _
         (D.PARENT_ID <> ParentID) Or _
         (D.PIG_TYPE <> PigType) Then
         
         D.PART_ITEM_ID = -1
         D.PART_TYPE = PartTypeID
         D.PIG_FLAG = PigFlag
         D.PIG_TYPE = PigType
         D.SPECIFIC_FLAG = SpecificFlag
         D.PARENT_ID = ParentID
         Call D.QueryData(1, Rs, itemcount)
      End If
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CPartItem
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC)
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.Add(TempData, TempData.PART_NO & "-" & TempData.PIG_TYPE)
         ElseIf KeyType = 3 Then
            Call SetPartKey(TempData, Cl, Trim(TempData.PIG_TYPE & "-" & TempData.PIG_FLAG & "-" & TempData.PART_NO & "-" & TempData.PART_DESC))
         End If
      End If
            
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Not Rs.BOF Then
      Rs.MoveFirst
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub SetPartKey(Cl As CPartItem, TempColl As Collection, Key As String)
On Error Resume Next
      Call TempColl.Add(Cl, Key)
End Sub
Public Sub LoadPigItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional PigType As String = "")
On Error GoTo ErrorHandler
Dim D As CPartItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartItem
Dim I As Long

   Set D = New CPartItem
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.PART_TYPE = -1
   D.PIG_FLAG = "Y"
   D.PIG_TYPE = PigType
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CPartItem
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC)
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_NO & "-" & TempData.PIG_TYPE)
'''debug.print TempData.PART_NO & "-" & TempData.PIG_TYPE
      End If
            
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadUnit(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUnit
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUnit
Dim I As Long

   Set D = New CUnit
   Set Rs = New ADODB.Recordset
   
   D.UNIT_ID = -1
   D.UNIT_NAME = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUnit
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.UNIT_NAME)
         C.ItemData(I) = TempData.UNIT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCnDnReason(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCnDnReason
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCnDnReason
Dim I As Long

   Set D = New CCnDnReason
   Set Rs = New ADODB.Recordset
   
   D.REASON_ID = -1
   D.REASON_NAME = ""
   D.REASON_NO = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCnDnReason
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.REASON_NAME)
         C.ItemData(I) = TempData.REASON_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.REASON_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBatch(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CBatch
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBatch
Dim I As Long

   Set D = New CBatch
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("BATCH_ID", -1)
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If glbUser.SIMULATE_FLAG = "N" Then
      Set Rs = Nothing
      Set D = Nothing
      
      Exit Sub
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBatch
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("BATCH_NO"))
         C.ItemData(I) = TempData.GetFieldValue("BATCH_ID")
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GetFieldValue("BATCH_ID"))))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExpenseType(C As ComboBox, Optional Cl As Collection = Nothing, Optional BuyFlag As String = "", Optional DeplicateFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExpenseType
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExpenseType
Dim I As Long

   Set D = New CExpenseType
   Set Rs = New ADODB.Recordset
   
   D.EXPENSE_TYPE_ID = -1
   D.BUY_FLAG = BuyFlag
   D.DEPLICATE_FLAG = DeplicateFlag
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExpenseType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.EXPENSE_TYPE_NAME)
         C.ItemData(I) = TempData.EXPENSE_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.EXPENSE_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice1(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.EXPORT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.UNIT_NAME)
'         C.ItemData(i) = TempData.UNIT_ID
      End If
            
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.HOUSE_ID) & "-" & Trim(TempData.PIG_ID) & "-" & Trim(TempData.PART_GROUP_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumFeedUsedAmountYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional BatchID As Long = -1, Optional PigType As String = "", Optional IntakeFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.EXPORT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.HOUSE_ID = LocationID
   D.PIG_FLAG = "N"
   D.BATCH_ID = BatchID
   D.TO_PIG_TYPE = PigType
   D.INTAKE_FLAG = IntakeFlag
   Call D.QueryData(45, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(45, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.UNIT_NAME)
'         C.ItemData(i) = TempData.UNIT_ID
      End If
            
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumFeedUsedAmountByPigYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional BatchID As Long = -1, Optional IntakeFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.EXPORT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.HOUSE_ID = LocationID
   D.PIG_FLAG = "N"
   D.BATCH_ID = BatchID
   D.INTAKE_FLAG = IntakeFlag
   Call D.QueryData(46, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(46, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.UNIT_NAME)
'         C.ItemData(i) = TempData.UNIT_ID
      End If
            
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
Dim Houses As Collection
Dim Hg As CHouseGroup

   Set Houses = New Collection
   Call LoadHouseGroup(Nothing, Houses, "N")
   
   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   Set Cl = Nothing
   Set Cl = New Collection
   
   For Each Hg In Houses
      D.EXPORT_ITEM_ID = -1
      D.FROM_DATE1 = FromDate
      D.TO_DATE1 = ToDate
      D.HOUSE_GROUP_ID = Hg.HOUSE_GROUP_ID
      D.COMMIT_FLAG1 = CommitFlag
      Call D.QueryData(5, Rs, itemcount)
   
      While Not Rs.EOF
         I = I + 1
         Set TempData = New CExportItem
         Call TempData.PopulateFromRS(5, Rs)
                     
         If Not (Cl Is Nothing) Then
            '''debug.print Trim(TempData.PART_TYPE_ID) & "-" & Trim(Hg.HOUSE_GROUP_ID) & "    :     " & TempData.EXPORT_TOTAL_PRICE
            Call Cl.Add(TempData, Trim(TempData.PART_TYPE_ID) & "-" & Trim(Hg.HOUSE_GROUP_ID))
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
   Next Hg
      
   Set D = Nothing
   Set Rs = Nothing
   Set Houses = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice3(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional HouseGroupID As Long, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
Dim Houses As Collection
Dim Hg As CLocation

   Set Houses = New Collection
   Call LoadLocationEx(Nothing, Houses, HouseGroupID)
   
   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   Set Cl = Nothing
   Set Cl = New Collection
   
   For Each Hg In Houses
      D.EXPORT_ITEM_ID = -1
      D.FROM_DATE1 = FromDate
      D.TO_DATE1 = ToDate
      D.HOUSE_ID1 = Hg.LOCATION_ID
      D.COMMIT_FLAG1 = CommitFlag
      Call D.QueryData(6, Rs, itemcount)

      While Not Rs.EOF
         I = I + 1
         Set TempData = New CExportItem
         Call TempData.PopulateFromRS(6, Rs)
                     
         If Not (Cl Is Nothing) Then
            Call Cl.Add(TempData, Trim(TempData.PART_TYPE_ID) & "-" & Trim(Hg.LOCATION_ID))
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
   Next Hg
      
   Set D = Nothing
   Set Rs = Nothing
   Set Houses = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice4(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(7, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)) & "-" & Trim(Str(TempData.HOUSE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice5(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(8, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)) & "-" & Trim(Str(TempData.HOUSE_ID)) & "-" & Trim(Str(TempData.PIG_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice6(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(10, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PART_NO) & "-" & Trim(TempData.PIG_TYPE) & "-" & Trim(Str(TempData.LOCATION_ID)))
'''debug.print Trim(TempData.PART_NO) & "-" & Trim(TempData.PIG_TYPE) & "-" & Trim(Str(TempData.LOCATION_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice7(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   '==
   FromDate1 = DateToStringIntLow(-2)
   ToDate1 = DateToStringIntLow(FromDate)
   '==
   
   D.FROM_DATE = InternalDateToDate(FromDate1)
   D.TO_DATE = InternalDateToDate(ToDate1)
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(11, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice8(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(12, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)) & "-" & Trim(Str(TempData.DOCUMENT_TYPE)) & "-" & Trim(TempData.SALE_FLAG))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice9(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(16, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(16, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PIG_ID) & "-" & Trim(TempData.PIG_TYPE) & "-" & Trim(Str(TempData.PART_TYPE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigSellAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional BillSubType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   D.DOCUMENT_TYPE = -1
   D.DocTypeSet = DocType2Set(BillSubType)
   Call D.QueryData(26, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(26, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PIG_STATUS)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice10(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(19, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(19, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.LOCATION_ID) & "-" & Trim(TempData.PIG_STATUS))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice12(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.LOCATION_ID = LocationID
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(21, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(21, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PART_ITEM_ID) & "-" & Trim(TempData.PIG_STATUS))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice11(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional StatusGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.STATUS_GROUP_ID = StatusGroupID
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(20, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(20, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.LOCATION_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice13(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional StatusGroupID As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.LOCATION_ID = LocationID
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.STATUS_GROUP_ID = StatusGroupID
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(22, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(22, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportPrice14(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PART_NO As String = "", Optional PART_GROUP_ID As Long = -1, Optional OrderType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.EXPORT_ITEM_ID = -1
   D.PIG_FLAG = "N"
   D.PART_NO = PART_NO
   D.PART_GROUP_ID = PART_GROUP_ID
   D.OrderBy = 12
   D.OrderType = OrderType
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMixPatial(Cl_Mix As Collection, Optional Cl_Im As Collection = Nothing, Optional Cl_Ex As Collection = Nothing)
On Error GoTo ErrorHandler
Dim tempEX As CExportItem
Dim TempIm As CExportItem



'Dim D As CExportItem
'Dim ItemCount As Long
'Dim Rs As ADODB.Recordset

'Dim I As Long
'
'   Set D = New CExportItem
'   Set Rs = New ADODB.Recordset
'
'   D.FROM_DATE = TempFromDate
'   D.TO_DATE = TempToDate
'   D.EXPORT_ITEM_ID = -1
'   D.PIG_FLAG = "N"
'   D.PART_NO = PART_NO
'   D.PART_GROUP_ID = PART_GROUP_ID
'   D.OrderBy = 12
'   D.OrderType = OrderType
'   Call D.QueryData(1, Rs, ItemCount)
'
'   If Not (C Is Nothing) Then
'      C.Clear
'      I = 0
'      C.AddItem ("")
'   End If
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
'   While Not Rs.EOF
'      I = I + 1
'      Set TempData = New CExportItem
'      Call TempData.PopulateFromRS(1, Rs)
'
'      If Not (C Is Nothing) Then
'      End If
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.Add(TempData)
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'
'   Set Rs = Nothing
'   Set D = Nothing
'   Exit Sub
'
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigStatusAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional StatusGroupID As Long = -1, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional HouseGroupID As Long = -1, Optional DocTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.LOCATION_ID = LocationID
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.STATUS_GROUP_ID = StatusGroupID
   D.COMMIT_FLAG = CommitFlag
   D.DOCUMENT_TYPE = DocumentType
   D.HOUSE_GROUP_ID = HouseGroupID
   D.DocTypeSet = DocTypeSet
   D.OrderBy = 1
   
   Call D.QueryData(24, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(24, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPrice1(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PART_NO) & "-" & Trim(TempData.PIG_TYPE) & "-" & Trim(Str(TempData.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBuyFeedYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(29, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(29, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBuyFeedYYYYMM2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional BatchID As Long = -1, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(30, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(30, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM2)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBuyFeedIntakeTypeYYYYMM2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional BatchID As Long = -1, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(32, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(32, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM2 & "-" & TempData.INTAKE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapDistinctPigHouseFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional ColParam As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional PigType As String = "", Optional DocumentType As Long = -1, Optional DocumentCat As Long = -1, Optional ParentFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.PIG_TYPE = PigType
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_CATEGORY = DocumentCat
   D.PARENT_FLAG = ParentFlag
   If Not (ColParam Is Nothing) Then
      D.OrderBy = ColParam("ORDER_BY")
      D.OrderType = ColParam("ORDER_TYPE")
   End If
   Call D.QueryData(10, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.TO_HOUSE_ID & "-" & TempData.PIG_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapDistinctPigFromToHouseFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional ColParam As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional PigType As String = "", Optional DocumentType As Long = -1, Optional DocumentCat As Long = -1, Optional ParentFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.PIG_TYPE = PigType
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_CATEGORY = DocumentCat
   D.PARENT_FLAG = ParentFlag
   If Not (ColParam Is Nothing) Then
      D.OrderBy = ColParam("ORDER_BY")
      D.OrderType = ColParam("ORDER_TYPE")
   End If
   Call D.QueryData(12, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.TO_HOUSE_ID & "-" & TempData.PIG_ID & "-" & TempData.TO_PIG_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapDistinctPigHouse(C As ComboBox, Optional Cl As Collection = Nothing, Optional ColParam As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional PigType As String = "", Optional DocumentType As Long = -1, Optional DocumentCat As Long = -1, Optional ParentFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.PIG_TYPE = PigType
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_CATEGORY = DocumentCat
   D.PARENT_FLAG = ParentFlag
   If Not (ColParam Is Nothing) Then
      D.OrderBy = ColParam("ORDER_BY")
      D.OrderType = ColParam("ORDER_TYPE")
   End If
   Call D.QueryData(5, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.PIG_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapDistinctPigStatus(C As ComboBox, Optional Cl As Collection = Nothing, Optional ColParam As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional PigType As String = "")
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.HOUSE_SALE_FLAG = "Y"
   D.PIG_TYPE = PigType
   If Not (ColParam Is Nothing) Then
      D.OrderBy = ColParam("ORDER_BY")
      D.OrderType = ColParam("ORDER_TYPE")
   End If
   Call D.QueryData(6, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_STATUS & "-" & TempData.PIG_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMovementLocation(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional YearID As Long = -1, Optional WeekNo As String, Optional PigType As String, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
Static Locations As Collection
Dim Lc As CLocation

   If Locations Is Nothing Then
      Set Locations = New Collection
      Call LoadLocation(Nothing, Locations, -1, "")
   End If
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.OrderBy = 1
   D.YEAR_SEQ_ID = YearID
   D.WEEK_NO = WeekNo
   D.PIG_TYPE = PigType
   D.BATCH_ID = BatchID
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Lc = Locations(Trim(Str(TempData.FROM_HOUSE_ID)))
      If Not (Cl Is Nothing) Then
         Call Cl.Add(Lc, Trim(Str(Lc.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Locations = Nothing
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExpRatioLocation(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExpenseRatio
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExpenseRatio
Dim I As Long
Static Locations As Collection
Dim Lc As CLocation

   If Locations Is Nothing Then
      Set Locations = New Collection
      Call LoadLocation(Nothing, Locations, 1, "")
   End If
   
   Set D = New CExpenseRatio
   Set Rs = New ADODB.Recordset
   
   D.EXPENSE_RATIO_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExpenseRatio
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Lc = Locations(Trim(Str(TempData.LOCATION_ID)))
      If Not (Cl Is Nothing) Then
         Call Cl.Add(Lc, Trim(Str(Lc.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapitalBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_ID = PigID
   D.OrderBy = 1
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
'   ''debug.print LocationID & "-" & PigID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE & "-" & TempData.CAPITAL_AMOUNT
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapitalBalanceEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional DocumentCat As Long = -1, Optional OutFlag As Long = 1, Optional TxType As String = "", Optional ReplaceFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_CATEGORY = DocumentCat
   D.PIG_ID = PigID
   D.TX_TYPE = TxType
   D.REPLACE_FLAG = ReplaceFlag
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(13, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(13, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapitalBalanceExFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional DocumentCat As Long = -1, Optional OutFlag As Long = 1, Optional TxType As String = "", Optional ReplaceFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_CATEGORY = DocumentCat
   D.PIG_ID = PigID
   D.TX_TYPE = TxType
   D.REPLACE_FLAG = ReplaceFlag
   D.OrderBy = 1
   Call D.QueryData(17, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(17, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.TO_HOUSE_ID & "-" & TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapitalBalanceExFromToPigFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional DocumentCat As Long = -1, Optional OutFlag As Long = 1, Optional TxType As String = "", Optional ReplaceFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_CATEGORY = DocumentCat
   D.PIG_ID = PigID
   D.TX_TYPE = TxType
   D.REPLACE_FLAG = ReplaceFlag
   D.OrderBy = 1
   Call D.QueryData(18, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.TO_HOUSE_ID & "-" & TempData.PIG_ID & "-" & TempData.TO_PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInitialCapitalBalance1(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long
Dim S As CMovementItemSearch1

Dim TempColl As Collection
Dim Ba As CBalanceAccum



   Set TempColl = New Collection
   Set Ba = New CBalanceAccum
   
   Call LoadInitialPigBalance(Nothing, TempColl, -1, ToDate, , , BatchID)
   
   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_ID = PigID
   D.DOCUMENT_CATEGORY = DocumentCat
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(11, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set S = New CMovementItemSearch1
         S.HOUSE_ID = TempData.FROM_HOUSE_ID
         S.PART_ITEM_ID = TempData.PART_ITEM_ID
         S.PIG_ID = TempData.PIG_ID
         S.CAPITAL_AMOUNT = TempData.CAPITAL_AMOUNT
'If S.PIG_ID & "-" & S.LOCATION_ID = "9341-278" Then
   '''debug.print
'End If
         Call Cl.Add(S, S.GetKey1)
         
         Set S = Nothing
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInitialCapitalBalance2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long
Dim S As CMovementItemSearch1

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_ID = PigID
   D.DOCUMENT_CATEGORY = DocumentCat
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(12, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set S = New CMovementItemSearch1
         S.HOUSE_ID = TempData.FROM_HOUSE_ID
         S.EXPENSE_TYPE = TempData.EXPENSE_TYPE
         S.PIG_ID = TempData.PIG_ID
         S.CAPITAL_AMOUNT = TempData.CAPITAL_AMOUNT
         
         Call Cl.Add(S, S.GetKey3)
         
         Set S = Nothing
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInitialPigBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim TempData As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim I As Long
Dim S As CImportItem

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset
   
   D.BALANCE_ACCUM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
'      If TempData.PART_ITEM_ID = 13862 And TempData.LOCATION_ID = 504 Then
'         ''debug.print
'      End If
      
      If Not (Cl Is Nothing) Then
         Set S = New CImportItem
         S.LOCATION_ID = TempData.LOCATION_ID
         S.PART_ITEM_ID = TempData.PART_ITEM_ID
'         S.PART_DESC = TempData.PART_DESC
'         S.LOCATION_NAME = TempData.LOCATION_NAME
         
         S.CURRENT_AMOUNT = TempData.BALANCE_AMOUNT
'If S.PART_ITEM_ID & "-" & S.LOCATION_ID = "8897-255" Then
''   ''debug.print
'End If
         Call Cl.Add(S, S.PART_ITEM_ID & "-" & S.LOCATION_ID)
         
'         If S.PART_ITEM_ID = 13749 And S.LOCATION_ID = 446 Then
'            ''debug.print S.CURRENT_AMOUNT
'         End If
         
         Set S = Nothing
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumCapitalBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional SaleFlag As String = "", Optional HouseGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.SALE_FLAG = SaleFlag
   D.HOUSE_GROUP_ID = HouseGroupID
   D.OrderBy = 1
   Call D.QueryData(7, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumCapitalSellBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional StatusGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.DOCUMENT_CATEGORY = 1
   D.DocTypeSet = "(10, 13) "
   D.STATUS_GROUP_ID = StatusGroupID
   D.OrderBy = 1
   Call D.QueryData(7, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(7, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumCapitalSellBalanceByStatus(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional StatusGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.DOCUMENT_CATEGORY = 1
   'D.DocTypeSet = "(10, 13) "
   'D.DOCUMENT_TYPE = 7              'เอกสาร H ใบโอนเข้าเรือนขาย
   D.DocTypeSet = "(7) "
   D.TX_TYPE = "E"
   D.STATUS_GROUP_ID = StatusGroupID
   D.OrderBy = 1
   Call D.QueryData(21, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(21, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE & "-" & TempData.PIG_STATUS)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalCapitalSellBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional StatusID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.DOCUMENT_CATEGORY = 1
   D.DocTypeSet = "(10, 13)"
   D.OrderBy = 1
   Call D.QueryData(10, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(10, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT

      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
'''debug.print TempData.PIG_STATUS
         Call Cl.Add(TempData, Trim(Str(TempData.PIG_STATUS)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalSellCapitalYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = -1
   D.PIG_STATUS = -1
   D.DOCUMENT_CATEGORY = 1
   D.DocTypeSet = "(10, 13)"
   D.OrderBy = 1
   Call D.QueryData(19, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(19, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT

      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalLossCapitalYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CLossItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLossItem
Dim I As Long

   Set D = New CLossItem
   Set Rs = New ADODB.Recordset
   
   D.LOSS_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = -1
   D.PIG_STATUS = -1
   D.DOCUMENT_CATEGORY = 1
'   D.DocTypeSet = "(10, 13)"
   D.OrderBy = 1
   Call D.QueryData(19, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLossItem
      Call TempData.PopulateFromRS(19, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT

      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTotalLossCapital(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional ParentFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CLossItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLossItem
Dim I As Long

   Set D = New CLossItem
   Set Rs = New ADODB.Recordset
   
   D.LOSS_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = -1
   D.PIG_STATUS = -1
   D.DOCUMENT_CATEGORY = 1
   D.PARENT_FLAG = ParentFlag
'   D.DocTypeSet = "(10, 13)"
   D.OrderBy = 1
   Call D.QueryData(20, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLossItem
      Call TempData.PopulateFromRS(20, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT

      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalCapitalSellBalancePigStatus(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional StatusID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.DOCUMENT_CATEGORY = 1
   D.DocTypeSet = "(10, 13)"
   D.OrderBy = 1
   Call D.QueryData(16, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(16, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PIG_STATUS)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTotalCapitalSellBalancePigStatusAgeCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional DocTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.DOCUMENT_CATEGORY = 1
   D.DocTypeSet = DocTypeSet
   D.OrderBy = 1
   Call D.QueryData(25, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(25, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_STATUS & "-" & TempData.AGE_CODE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTotalCapitalSellBalancePigStatusPigID(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional DocTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long
   
   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.DOCUMENT_CATEGORY = 1
   D.DocTypeSet = DocTypeSet
   D.OrderBy = 1
   
   Call D.QueryData(26, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(26, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_STATUS & "-" & TempData.PIG_AGE & "-" & TempData.PIG_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumRevenueItemCost(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CRevenueCostItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRevenueCostItem
Dim I As Long

   Set D = New CRevenueCostItem
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CRevenueCostItem
      Call TempData.PopulateFromRS(2, Rs)

      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumPigStatusSellBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional StatusID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.DOCUMENT_CATEGORY = 1
   D.DOCUMENT_TYPE = 10
   D.OrderBy = 1
   Call D.QueryData(9, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(9, Rs)
      TempData.CAPITAL_AMOUNT = -1 * TempData.CAPITAL_AMOUNT
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_STATUS & "-" & TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumPigHouseBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional PigType As String = "")
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.PIG_TYPE = PigType
'   D.OrderBy = 1
   Call D.QueryData(8, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumHouseBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_CATEGORY = DocumentCat
'   D.OrderBy = 1
   Call D.QueryData(15, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumPigStatusBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional PigType As String = "")
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.PIG_TYPE = PigType
   D.OrderBy = 1
   Call D.QueryData(9, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_STATUS & "-" & TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapitalMovement(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_ID = PigID
   D.OrderBy = 1
   Call D.QueryData(6, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.GetKey1)
'''debug.print TempData.GetKey1
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCapitalMovementEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_ID = PigID
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(14, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.PIG_ID & "-" & TempData.GetKey1)
'''debug.print TempData.GetKey1
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMovementPig(C As ComboBox, Optional Cl As Collection = Nothing, Optional Params As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional ParentFlag As String = "", Optional YearID As Long = -1, Optional WeekNo As String, Optional PigType As String, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
Static Partitems As Collection
Dim Pi As CPartItem

   If Partitems Is Nothing Then
      Set Partitems = New Collection
      Call LoadPartItem(Nothing, Partitems, , "Y")
   End If

   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.PARENT_FLAG = ParentFlag
   D.PIG_FLAG = "Y"
   D.YEAR_SEQ_ID = YearID
   D.WEEK_NO = WeekNo
   D.BATCH_ID = BatchID
   If Not (Params Is Nothing) Then
      D.OrderBy = Params("ORDER_BY")
      D.OrderType = Params("ORDER_TYPE")
   Else
      D.OrderBy = 1
   End If
   D.OrderType = 1
   D.PIG_TYPE = PigType
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set Pi = Partitems(Trim(Str(TempData.PIG_ID)))
         Call Cl.Add(Pi, Trim(Str(Pi.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Partitems = Nothing
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMovementSellPig(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional StatusID = -1, Optional ParentFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
Static Partitems As Collection
Dim Pi As CPartItem

   If Partitems Is Nothing Then
      Set Partitems = New Collection
      Call LoadPartItem(Nothing, Partitems, , "Y")
   End If
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.DOCUMENT_CATEGORY = 1
   D.DocTypeSet = "(10, 13)"
   D.PIG_STATUS = StatusID
   D.PARENT_FLAG = ParentFlag
   D.OrderBy = 1
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set Pi = Partitems(Trim(Str(TempData.PIG_ID)))
         Call Cl.Add(Pi, Trim(Str(Pi.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMovementSellPigByStatus(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional StatusID = -1, Optional ParentFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
Dim Pi As CPartItem
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.DOCUMENT_CATEGORY = 1
   'D.DocTypeSet = "(10, 13)"
   D.DOCUMENT_TYPE = 7
   D.PIG_STATUS = StatusID
   D.PARENT_FLAG = ParentFlag
   D.OrderBy = 1
   Call D.QueryData(14, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PRODUCT_STATUS_NO & "-" & TempData.PART_NO)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMovementSellPigStatus(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional StatusID = -1)
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
Static Partitems As Collection
Dim Pi As CPartItem

   If Partitems Is Nothing Then
      Set Partitems = New Collection
      Call LoadPartItem(Nothing, Partitems, , "Y")
   End If
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.CAPITAL_MOVEMENT_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.DOCUMENT_CATEGORY = 1
   D.DOCUMENT_TYPE = 10
   D.PIG_STATUS = StatusID
   D.OrderBy = 1
   Call D.QueryData(6, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PIG_STATUS)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPrice2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   '==
   FromDate1 = DateToStringIntLow(-2)
   ToDate1 = DateToStringIntLow(FromDate)
   '==
   
   D.FROM_DATE = InternalDateToDate(FromDate1)
   D.TO_DATE = InternalDateToDate(ToDate1)
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(5, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPrice3(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(6, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)) & "-" & Trim(Str(TempData.DOCUMENT_TYPE)) & "-" & Trim(TempData.SALE_FLAG))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPrice4(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(11, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)) & "-" & Trim(Str(TempData.DOCUMENT_TYPE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPrice5(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(13, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(13, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)) & "-" & Trim(Str(TempData.DOCUMENT_TYPE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadImportPrice6(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(14, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)) & "-" & Trim(Str(TempData.LOCATION_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadImportPrice7(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PART_NO As String = "", Optional PART_GROUP_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim tempCl As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PART_NO = PART_NO
   D.PART_GROUP = PART_GROUP_ID
   D.OrderBy = 7
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set tempCl = GetObject("CImport", Cl, TempData.PART_NO, False)
         If tempCl Is Nothing Then
            Call Cl.Add(TempData, TempData.PART_NO)
            Set tempCl = GetObject("CImport", Cl, TempData.PART_NO, False)
            Call tempCl.m_Import.Add(TempData)
         Else
            Call tempCl.m_Import.Add(TempData)
         End If
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumBalanceAccum(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = 1
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctPartLocationBA(C As ComboBox, mcolParam As Collection, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigFlag As String = "", Optional LocationType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = PigFlag
   D.LOCATION_TYPE = LocationType
   D.OrderBy = mcolParam("ORDER_BY")
   D.OrderType = mcolParam("ORDER_TYPE")
   D.LOCATION_ID = mcolParam("LOCATION_ID")
   D.PART_TYPE = mcolParam("PART_TYPE")
   Call D.QueryData(8, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumBalanceAccum2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = -1
   D.TO_DATE1 = ToDate
   D.OrderBy = 1
   'Call D.QueryData(5, Rs, ItemCount) 'เปลี่ยนจาก ind 5 เป็น 20
   Call D.QueryData(20, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      'Call TempData.PopulateFromRS(5, Rs)
      Call TempData.PopulateFromRS(20, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
'If TempData.PART_ITEM_ID = 7366 Then
'''debug.print
'End If
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumBalanceAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional HouseGroup As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE1 = ToDate
   D.HOUSE_GROUP_ID = HouseGroup
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   D.LOCATION_ID = LocationID
   Call D.QueryData(11, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumPigBalanceAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional HouseGroup As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE1 = ToDate
   D.HOUSE_GROUP_ID = HouseGroup
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   D.LOCATION_ID = LocationID
   Call D.QueryData(12, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ParentFlag As String = "", Optional SaleFlag As String = "", Optional HouseGroupID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_STATUS = StatusID
   D.PARENT_FLAG = ParentFlag
   D.PIG_FLAG = "Y"
   D.SALE_FLAG = SaleFlag
   D.HOUSE_SALE_FLAG = SaleFlag
   D.HOUSE_GROUP_ID = HouseGroupID
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(15, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigTypeImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ParentFlag As String = "", Optional SaleFlag As String = "", Optional HouseGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_STATUS = StatusID
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   Call D.QueryData(24, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(24, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_TYPE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigTypeDocTypeImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ParentFlag As String = "", Optional SaleFlag As String = "", Optional HouseGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_STATUS = StatusID
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   Call D.QueryData(25, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(25, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCUMENT_TYPE & "-" & TempData.PIG_TYPE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigBuyAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_STATUS = -1
   D.PIG_FLAG = "Y"
   D.DOCUMENT_TYPE = 11
   D.OrderBy = 1
   Call D.QueryData(15, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigExportByHouse(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional DocumentType As Long = -1, Optional ParentFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_STATUS = StatusID
   D.PIG_FLAG = "Y"
   D.DOCUMENT_TYPE = DocumentType
   D.PARENT_FLAG = ParentFlag
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(34, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(34, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigImportByHouse(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ParentFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_STATUS = StatusID
   D.PIG_FLAG = "Y"
   D.PARENT_FLAG = ParentFlag
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(22, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(22, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigHouseExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ParentFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_STATUS = StatusID
   D.PIG_FLAG = "Y"
   D.PARENT_FLAG = ParentFlag
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(30, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(30, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.LOCATION_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigHouseImportAmountFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional ParentFlag As String = "", Optional DocumentType As Long = -1, Optional ReplaceFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PARENT_FLAG = ParentFlag
   D.DOCUMENT_TYPE = DocumentType
   D.REPLACE_FLAG = ReplaceFlag
   D.OrderBy = 1
   Call D.QueryData(18, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.LOCATION_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigHouseImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional ParentFlag As String = "", Optional DocumentType As Long = -1, Optional ReplaceFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PARENT_FLAG = ParentFlag
   D.DOCUMENT_TYPE = DocumentType
   D.REPLACE_FLAG = ReplaceFlag
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(18, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.LOCATION_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigHouseImportCmAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional ParentFlag As String = "", Optional DocumentType As Long = -1, Optional ReplaceFlag As String = "", Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.FROM_HOUSE_ID = LocationID
   D.PARENT_FLAG = ParentFlag
   D.DOCUMENT_TYPE = DocumentType
   D.REPLACE_FLAG = ReplaceFlag
   D.TX_TYPE = TxType
   D.OrderBy = 1
   Call D.QueryData(11, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.FROM_HOUSE_ID & "-" & TempData.TO_HOUSE_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigFromToHouseImportCmAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional ParentFlag As String = "", Optional DocumentType As Long = -1, Optional ReplaceFlag As String = "", Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.FROM_HOUSE_ID = LocationID
   D.PARENT_FLAG = ParentFlag
   D.DOCUMENT_TYPE = DocumentType
   D.REPLACE_FLAG = ReplaceFlag
   D.TX_TYPE = TxType
   D.OrderBy = 1
   Call D.QueryData(13, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(13, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.TO_PIG_ID & "-" & TempData.FROM_HOUSE_ID & "-" & TempData.TO_HOUSE_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigStatusImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   Call D.QueryData(21, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(21, Rs)

      If Not (C Is Nothing) Then
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PIG_STATUS)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigHouseStatusImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   Call D.QueryData(20, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(20, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.PIG_STATUS)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.DOCUMENT_TYPE = DocumentType
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "N"
   D.OrderBy = 1
'D.PART_ITEM_ID = 10443
   Call D.QueryData(15, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ParentFlag As String = "", Optional SaleFlag As String = "", Optional DocumentType As Long = -1, Optional HouseGroupID As Long = -1, Optional BatchID As Long = -1, Optional DocTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
   
   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PIG_STATUS = StatusID
   D.DOCUMENT_TYPE = DocumentType
   D.PARENT_FLAG = ParentFlag
   D.HOUSE_SALE_FLAG = SaleFlag
   D.HOUSE_GROUP_ID = HouseGroupID
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   D.DocTypeSet = DocTypeSet
   Call D.QueryData(23, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(23, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPigFeedAmountYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "N"
   D.DOCUMENT_TYPE = 2
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(42, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(42, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigFeedAmountYYYYMM2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.HOUSE_ID = LocationID
   D.PIG_FLAG = "N"
   D.DOCUMENT_TYPE = 2
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(50, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(50, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PIG_ID & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigFeedAmountByFeed(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional BatchID As Long = -1, Optional HouseGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.HOUSE_ID = LocationID
   D.TO_HOUSE_GROUP_ID = HouseGroup
   D.PIG_FLAG = "N"
   D.DOCUMENT_TYPE = 2
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(48, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(48, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
'''debug.print TempData.PART_ITEM_ID & " " & TempData.EXPORT_AMOUNT
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigFeedAmountByFeedPig(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional BatchID As Long = -1, Optional HouseGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.HOUSE_ID = LocationID
   D.TO_HOUSE_GROUP_ID = HouseGroup
   D.PIG_FLAG = "N"
   D.DOCUMENT_TYPE = 2
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(52, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(52, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PART_ITEM_ID)
'''debug.print TempData.PART_ITEM_ID & " " & TempData.EXPORT_AMOUNT
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPigFeedAmountByFeedPigByPigAge(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional BatchID As Long = -1, Optional HouseGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.HOUSE_ID = LocationID
   D.TO_HOUSE_GROUP_ID = HouseGroup
   D.PIG_FLAG = "N"
   D.DOCUMENT_TYPE = 2
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(55, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(55, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.PIG_AGE)
'''debug.print TempData.PART_ITEM_ID & " " & TempData.EXPORT_AMOUNT
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigTypeExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ParentFlag As String = "", Optional SaleFlag As String = "", Optional DocumentType As Long = -1, Optional HouseGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PIG_STATUS = StatusID
   D.OrderBy = 1
   Call D.QueryData(37, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(37, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigTypeDocTypeExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ParentFlag As String = "", Optional SaleFlag As String = "", Optional DocumentType As Long = -1, Optional HouseGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PIG_STATUS = StatusID
   D.OrderBy = 1
   Call D.QueryData(38, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(38, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCUMENT_TYPE & "-" & TempData.PIG_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigHouseExportAmountEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional ParentFlag As String = "", Optional ReplaceFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.DOCUMENT_TYPE = DocumentType
   D.PARENT_FLAG = ParentFlag
   D.REPLACE_FLAG = ReplaceFlag
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(30, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(30, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartHouseExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.PIG_FLAG = "N"
   D.DOCUMENT_TYPE = -1
   D.OrderBy = 1
   Call D.QueryData(30, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(30, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaymentByType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PaymentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPaymentItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPaymentItem
Dim I As Long

   Set D = New CPaymentItem
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PAYMENT_TYPE = PaymentType
   D.OrderBy = 1
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPaymentItem
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PAYMENT_TYPE & "-" & TempData.TX_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaymentByDocTypeSubType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PaymentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPaymentItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPaymentItem
Dim I As Long

   Set D = New CPaymentItem
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PAYMENT_TYPE = PaymentType
   D.OrderBy = 1
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPaymentItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCUMENT_TYPE & "-" & TempData.RECEIPT_TYPE & "-" & TempData.PAYMENT_TYPE & "-" & TempData.TX_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigStatusExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   Call D.QueryData(32, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(32, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PIG_STATUS)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigMonthlyAccumYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMonthlyAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMonthlyAccum
Dim I As Long

   Set D = New CMonthlyAccum
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   D.OrderType = 2 'ต้องเรียงจากมากไปน้อย
   'D.PART_ITEM_ID = 12641
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMonthlyAccum
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigHouseStatusExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   Call D.QueryData(31, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(31, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.PIG_STATUS)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSalePigExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ExceptionFlag As String = "", Optional CapitalMovementFlag As String = "", Optional StatusGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
'   D.DOCUMENT_TYPE = 10
   D.DocTypeSet = "(10, 13)"
   D.PIG_FLAG = "Y"
   D.CAPITAL_MOVE_FLAG = CapitalMovementFlag
   D.EXCEPTION_FLAG = ExceptionFlag
   D.PIG_STATUS = StatusID
   D.OrderBy = 1
    D.STATUS_GROUP_ID = StatusGroupID
   Call D.QueryData(23, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(23, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSalePigExportAmountByStatus(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1, Optional ExceptionFlag As String = "", Optional CapitalMovementFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
'   D.DOCUMENT_TYPE = 10
   'D.DocTypeSet = "(10, 13)"
   D.DOCUMENT_TYPE = 7
   D.PIG_FLAG = "Y"
   D.CAPITAL_MOVE_FLAG = CapitalMovementFlag
   D.EXCEPTION_FLAG = ExceptionFlag
   D.PIG_STATUS = StatusID
   D.OrderBy = 1
   Call D.QueryData(56, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(56, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PIG_STATUS & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSalePigStatusExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional StatusID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
'   D.DOCUMENT_TYPE = 10
   D.DocTypeSet = "(10, 13)"
   D.PIG_FLAG = "Y"
   D.PIG_STATUS = StatusID
   D.OrderBy = 1
   Call D.QueryData(32, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(32, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PIG_STATUS)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSalePartCustAmountPrice(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional CustomerType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.CUSTOMER_TYPE = CustomerType
   D.PIG_FLAG = "N"
   D.OrderBy = 1
   Call D.QueryData(10, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSTOMER_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSalePartCustAmountPrice_SLMKEY(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional CustomerType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.CUSTOMER_TYPE = CustomerType
 '  D.PIG_FLAG = "N"
   D.OrderBy = 1
   Call D.QueryData(10, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSTOMER_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.EMP_ID)
         ''debug.print TempData.CUSTOMER_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.EMP_ID
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSalePartCustAmountPig(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional CustomerType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
  ' D.CUSTOMER_TYPE = CustomerType
   D.PIG_FLAG = "Y"
'   D.OrderBy = 1
   Call D.QueryData(61, Rs, itemcount)
   
'      Ei1.EXPORT_ITEM_ID = -1
'      Ei1.FROM_DATE = mcolParam("FROM_DATE")
'      Ei1.TO_DATE = mcolParam("TO_DATE")
'      Ei1.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
''      Ei1.LOCATION_ID = mcolParam("HOUSE_ID")
''      Ei1.DOCUMENT_TYPE = -1 'mcolParam("BILL_SUBTYPE")
''      Ei1.DocTypeSet = DocType2Set(mcolParam("BILL_SUBTYPE"))
''      Ei1.PIG_STATUS = mcolParam("STATUS_ID")
'      Ei1.PIG_FLAG = "Y"
''      Ei1.OrderBy = mcolParam("ORDER_BY")
''      Ei1.OrderType = mcolParam("ORDER_TYPE")
'      Call Ei1.QueryData(57, Rs, iCount)
'
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(61, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUS_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.EMP_ID)
         ''debug.print TempData.CUS_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.EMP_ID
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSaleRevenueCustAmountPrice(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional CustomerType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.CUSTOMER_TYPE = CustomerType
   D.RevenueFlag = "U"
   D.OrderBy = 1
   Call D.QueryData(12, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSTOMER_ID & "-" & TempData.REVENUE_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSalePigStatusYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   
   Call D.QueryData(25, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(25, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PIG_STATUS & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigParentUseAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(29, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(29, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PIG_ID & "-" & TempData.HOUSE_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPigParentUseAmountEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(53, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(53, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         'Call Cl.Add(TempData, Trim(TempData.PIG_ID & "-" & TempData.HOUSE_ID))
         Call Cl.Add(TempData, Trim(Str(TempData.HOUSE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRevenueAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RevenueID As Long = -1, Optional BillSubType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.REVENUE_ID = RevenueID
   D.DocTypeSet = BillingDocType2Set(BillSubType)
   D.OrderBy = 1
   Call D.QueryData(7, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.REVENUE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExpenseAmountYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CROItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CROItem
Dim I As Long

   Set D = New CROItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CROItem
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM & "-" & TempData.EXPENSE_TYPE & "-" & TempData.EXPENSE_DESC)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalExpenseAmountYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CROItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CROItem
Dim I As Long

   Set D = New CROItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CROItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTotalExpenseAmountYYYYMM2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1, Optional DepreciationGoodFlag As String)
On Error GoTo ErrorHandler
Dim D As CROItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CROItem
Dim I As Long
   
   Set D = New CROItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   D.DEPRECIATION_GOOD_FLAG = DepreciationGoodFlag
   D.OrderBy = 1
   
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CROItem
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM2)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigStatusAmountYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(43, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(43, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
'If TempData.PIG_STATUS > 0 Then
'''debug.print
'End If
         Call Cl.Add(TempData, TempData.YYYYMM & "-" & TempData.PART_ITEM_ID & "-" & TempData.PIG_STATUS)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTotalSellPrice(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RevenueID As Long = -1, Optional BillSubType As Long = -1, Optional BatchID As Long = -1, Optional NonSaleFlag As String)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.REVENUE_ID = RevenueID
   D.BATCH_ID = BatchID
   D.DocTypeSet = BillingDocType2Set(BillSubType)
   D.OrderBy = 1
   D.NON_SALE_FLAG = NonSaleFlag
   Call D.QueryData(16, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(16, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalCashSellPig(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RevenueID As Long = -1, Optional BillSubType As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.REVENUE_ID = RevenueID
   D.DocTypeSet = BillingDocType2Set(BillSubType)
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(18, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM2)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalRevenuePrice(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RevenueID As Long = -1, Optional BillSubType As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.REVENUE_ID = RevenueID
   D.BATCH_ID = BatchID
   D.DocTypeSet = BillingDocType2Set(BillSubType)
   D.OrderBy = 1
   Call D.QueryData(17, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(17, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalRevenueCash(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RevenueID As Long = -1, Optional BillSubType As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.REVENUE_ID = RevenueID
   D.DocTypeSet = BillingDocType2Set(BillSubType)
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(19, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(19, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM2)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigParentUseAmountEx1(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(35, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(35, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_GROUP_ID & "-" & TempData.PIG_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigParentUseAmountEx2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   Call D.QueryData(36, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(36, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.HOUSE_ID & "-" & TempData.PART_GROUP_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigParentUseAmountEx3(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(39, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(39, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigPartUseAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional IntakeFlag As String = "", Optional HouseGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.INTAKE_FLAG = IntakeFlag
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
'   D.HOUSE_GROUP_ID = HouseGroupID
   D.TO_HOUSE_GROUP_ID = HouseGroupID
   D.OrderBy = 1
   Call D.QueryData(25, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(25, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.SALE_FLAG)
'''debug.print "0:" & TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.SALE_FLAG
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartExportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional DocumentType As Long = -1, Optional SaleFlag As String = "", Optional HouseGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
   
   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.PIG_FLAG = "N"
   D.SALE_FLAG = SaleFlag
   D.DOCUMENT_TYPE = DocumentType
   D.OrderBy = 1
   D.TO_HOUSE_GROUP_ID = HouseGroupID
   
   Call D.QueryData(23, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(23, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartSellAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.PIG_FLAG = "N"
   D.DocTypeSet = "(10, 13)"
   D.OrderBy = 1
   Call D.QueryData(23, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(23, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPig(C As ComboBox, Optional Cl As Collection = Nothing, Optional LocationID As Long = -1, Optional PigType As String = "", Optional BatchID As Long = -1, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
Dim EI As CExportItem
Dim TempPi As CPartItem

   If Partitems Is Nothing Then
      Set Partitems = New Collection
      Call LoadPartItem(Nothing, Partitems, , "Y", "", "")
   End If
   
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.PIG_FLAG = "Y"
   D.IMPORT_ITEM_ID = -1
   D.LOCATION_ID = LocationID
   D.PIG_TYPE = PigType
   D.OrderBy = 1
   D.BATCH_ID = BatchID
'D.PART_ITEM_ID = 14553               ' เดี่ยวลบทิ้ง
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(9, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Pi = Partitems(Trim(Str(TempData.PART_ITEM_ID)))
      If Not (Cl Is Nothing) Then
         Call Cl.Add(Pi, Trim(Str(Pi.PART_ITEM_ID)))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set EI = New CExportItem
   EI.PIG_FLAG = "Y"
   EI.EXPORT_ITEM_ID = -1
   EI.LOCATION_ID = LocationID
   EI.PIG_TYPE = PigType
   EI.FROM_DATE = FromDate
   EI.TO_DATE = ToDate
   EI.OrderBy = 1
   Call EI.QueryData(33, Rs, itemcount)
   Set EI = Nothing
   While Not Rs.EOF
      I = I + 1
      Set EI = New CExportItem
      Call EI.PopulateFromRS(33, Rs)
      
      Set TempPi = GetPartItem(Cl, Trim(Str(EI.PART_ITEM_ID)))
      If TempPi.PART_ITEM_ID <= 0 Then
         Set Pi = Partitems(Trim(Str(EI.PART_ITEM_ID)))
         If Not (Cl Is Nothing) Then
            Call Cl.Add(Pi, Str(Trim(Pi.PART_ITEM_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set EI = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExpenseRatio(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExpenseTypeID As Long, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CExpenseRatio
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExpenseRatio
Dim I As Long

   Set D = New CExpenseRatio
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = -1
   D.TO_DATE = -1
   D.EXPENSE_RATIO_ID = -1
   D.RO_ITEM_ID = ExpenseTypeID
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExpenseRatio
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadHouseExpenseRatio(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional DeplicateFlag As String = "")
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CExpenseRatio
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExpenseRatio
Dim I As Long

   Set D = New CExpenseRatio
   Set Rs = New ADODB.Recordset

   D.EXPENSE_RATIO_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DEPLICATE_FLAG = DeplicateFlag
   D.OrderBy = 1
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExpenseRatio
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExpExpenseRatio(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CExpenseRatio
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExpenseRatio
Dim I As Long

   Set D = New CExpenseRatio
   Set Rs = New ADODB.Recordset
   
   D.EXPENSE_RATIO_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = 1
   Call D.QueryData(5, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExpenseRatio
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.EXPENSE_TYPE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadHouseExpExpenseRatio(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional HouseId As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CExpenseRatio
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExpenseRatio
Dim I As Long

   Set D = New CExpenseRatio
   Set Rs = New ADODB.Recordset
   
   D.EXPENSE_RATIO_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = HouseId
   D.OrderBy = 1
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExpenseRatio
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPigEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional LocationID As Long = -1, Optional ToDate As Date = -1, Optional PigType As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
   
   If Partitems Is Nothing Then
      Set Partitems = New Collection
      Call LoadPartItem(Nothing, Partitems, , "Y")
   End If
   
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.PIG_FLAG = "Y"
   D.IMPORT_ITEM_ID = -1
   D.LOCATION_ID = LocationID
'D.PART_ITEM_ID = 14553
   D.PIG_TYPE = PigTypeToCode(PigType)
   D.OrderBy = 2
   D.OrderType = 2
   Call D.QueryData(9, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Pi = Partitems(Trim(Str(TempData.PART_ITEM_ID)))
      Pi.PIG_AGE = GetAge(Pi.PART_NO, ToDate)
      Pi.AGE_CODE = GetAgeCode(Pi.PIG_AGE)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(Pi, Str(Trim(Pi.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPigBirth(C As ComboBox, Optional Cl As Collection = Nothing, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
   
   If Partitems Is Nothing Then
      Set Partitems = New Collection
      Call LoadPartItem(Nothing, Partitems, , "Y")
   End If
   
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.PIG_FLAG = "Y"
   D.DOCUMENT_TYPE = 5
   D.IMPORT_ITEM_ID = -1
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(9, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Pi = Partitems(Trim(Str(TempData.PART_ITEM_ID)))
      If Not (Cl Is Nothing) Then
         Call Cl.Add(Pi, Trim(Str(Pi.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigBirthAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PigID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
      
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.IMPORT_ITEM_ID = -1
   D.PIG_FLAG = "Y"
   D.DOCUMENT_TYPE = 5              ' เอกสารการเกิดของงสุกร
   D.COMMIT_FLAG = CommitFlag
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PART_ITEM_ID = PigID
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(16, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(16, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, "1")
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadHousePigBirthAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PigID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
      
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.IMPORT_ITEM_ID = -1
   D.PIG_FLAG = "Y"
   D.DOCUMENT_TYPE = 5
   D.COMMIT_FLAG = CommitFlag
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PART_ITEM_ID = PigID
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(22, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(22, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadHousePigImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PigID As Long = -1, Optional DocumentType As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
      
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.IMPORT_ITEM_ID = -1
   D.PIG_FLAG = "Y"
   D.DOCUMENT_TYPE = DocumentType
   D.COMMIT_FLAG = CommitFlag
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PART_ITEM_ID = PigID
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(18, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
'''debug.print TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & " " & TempData.IMPORT_AMOUNT
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadHousePartImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
      
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.IMPORT_ITEM_ID = -1
   D.PIG_FLAG = "N"
   D.DOCUMENT_TYPE = -1
   D.COMMIT_FLAG = CommitFlag
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   Call D.QueryData(18, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadImportLocation(C As ComboBox, Optional Cl As Collection = Nothing, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Static Locations As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CLocation
   
   If Locations Is Nothing Then
      Set Locations = New Collection
      Call LoadLocation(Nothing, Locations, 1, "")
   End If
   
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.OrderBy = 1
   D.LOCATION_ID = LocationID
   Call D.QueryData(12, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Pi = Locations(Trim(Str(TempData.LOCATION_ID)))
      If Not (Cl Is Nothing) Then
         Call Cl.Add(Pi, Str(Trim(Pi.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartType As Long = -1, Optional LocationID As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional IntakeFlag As String = "", Optional PartNo As String = "", Optional OrderBy As Long = 1, Optional HouseGroupID As Long = -1)
On Error GoTo ErrorHandler
Static Partitems As Collection
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
Dim EI As CExportItem
Dim TempPi As CPartItem

   If Partitems Is Nothing Then
      Set Partitems = New Collection
      Call LoadPartItem(Nothing, Partitems, , "N")
   End If
   
   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.INTAKE_FLAG = IntakeFlag
   D.PIG_FLAG = "N"
   D.IMPORT_ITEM_ID = -1
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.COMMIT_FLAG = CommitFlag
   D.PART_NO = PartNo
   D.OrderBy = OrderBy
   D.HOUSE_GROUP_ID = HouseGroupID
   Call D.QueryData(9, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Pi = Partitems(Trim(Str(TempData.PART_ITEM_ID)))
'If TempData.PART_ITEM_ID = 7289 Then
'''debug.print
'End If
      If Not (Cl Is Nothing) Then
         Call Cl.Add(Pi, Trim(Str(Pi.PART_ITEM_ID)))
      End If
            
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set EI = New CExportItem
   EI.INTAKE_FLAG = IntakeFlag
   EI.PIG_FLAG = "N"
   EI.EXPORT_ITEM_ID = -1
   EI.LOCATION_ID = LocationID
   EI.PART_TYPE = PartType
   EI.OrderBy = OrderBy
   EI.HOUSE_GROUP_ID = HouseGroupID
   Call EI.QueryData(33, Rs, itemcount)
   Set EI = Nothing

   While Not Rs.EOF
      I = I + 1
      Set EI = New CExportItem
      Call EI.PopulateFromRS(33, Rs)
'If EI.PART_ITEM_ID = 12510 Then
''''debug.print
'End If

      Set TempPi = GetPartItem(Cl, Trim(Str(EI.PART_ITEM_ID)))
      If TempPi.PART_ITEM_ID <= 0 Then
         Set Pi = Partitems(Trim(Str(EI.PART_ITEM_ID)))
         If Not (Cl Is Nothing) Then
            Call Cl.Add(Pi, Str(Trim(Pi.PART_ITEM_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadImportPartItemEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartType As Long = -1, Optional LocationID As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional IntakeFlag As String = "", Optional PartNo As String = "")
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pi As CPartItem
Dim EI As CExportItem
Dim TempPi As CPartItem

   Set D = New CImportItem
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   
   D.INTAKE_FLAG = IntakeFlag
   D.PIG_FLAG = "N"
   D.IMPORT_ITEM_ID = -1
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.COMMIT_FLAG = CommitFlag
   D.PART_NO = PartNo
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = 4
   D.OrderType = 1
'D.PART_ITEM_ID = 10443
   Call D.QueryData(9, Rs1, itemcount)
'''debug.print "===="
'While Not (Rs1.EOF)
'   Call D.PopulateFromRS(9, Rs1)
'   ''debug.print D.PART_TYPE_NO & "|" & D.PART_NO
'   Rs1.MoveNext
'Wend
'Rs1.MoveFirst
   Set EI = New CExportItem
   EI.INTAKE_FLAG = IntakeFlag
   EI.PIG_FLAG = "N"
   EI.EXPORT_ITEM_ID = -1
   EI.LOCATION_ID = LocationID
   EI.PART_TYPE = PartType
   EI.FROM_DATE = FromDate
   EI.TO_DATE = ToDate
   EI.OrderBy = 4
   EI.OrderType = 1
   Call EI.QueryData(33, Rs2, itemcount)

'''debug.print "===="
'While Not (Rs2.EOF)
'   Call EI.PopulateFromRS(33, Rs2)
'   ''debug.print EI.PART_TYPE_NO & "|" & EI.PART_NO
'   Rs2.MoveNext
'Wend
'Rs2.MoveFirst
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
'''debug.print "===="
   While Not (Rs1.EOF And Rs2.EOF)
      Set Pi = GetMinPartItem1(Rs1, Rs2)
      If Pi.PART_NO = "153" Then
         ''debug.print
      End If
'''debug.print Pi.PART_ITEM_ID & "|" & Pi.PART_TYPE_NO & "|" & Pi.PART_NO
      Call Cl.Add(Pi, Trim(Str(Pi.PART_ITEM_ID)))
   Wend
   
   If Rs1.State = adStateOpen Then
      Call Rs1.Close
   End If
   Set Rs1 = Nothing
   
   If Rs2.State = adStateOpen Then
      Call Rs2.Close
   End If
   Set Rs2 = Nothing
   
   Set D = Nothing
   Set EI = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Function GetMinPartItem1(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset) As CPartItem
Dim Pi As CPartItem
Dim II As CImportItem
Dim EI As CExportItem
Dim O As Object

   Set II = New CImportItem
   Set EI = New CExportItem
   
   If Not Rs1.EOF Then
      Call II.PopulateFromRS(9, Rs1)
      
'      If Ii.PART_ITEM_ID = 12196 Then
'         ''debug.print
'      End If
   Else
      Set II = Nothing
   End If
   
   If Not Rs2.EOF Then
      Call EI.PopulateFromRS(33, Rs2)
'      If Ei.PART_ITEM_ID = 12196 Then
'         ''debug.print
'      End If
   Else
      Set EI = Nothing
   End If
   
   If II Is Nothing Then
      Set O = EI
      Call Rs2.MoveNext
   ElseIf EI Is Nothing Then
      Set O = II
      Call Rs1.MoveNext
   ElseIf II.PART_TYPE_NO < EI.PART_TYPE_NO Then
      Set O = II
      Call Rs1.MoveNext
   ElseIf II.PART_TYPE_NO > EI.PART_TYPE_NO Then
      Set O = EI
      Call Rs2.MoveNext
   Else
      If II.PART_NO < EI.PART_NO Then
         Set O = II
         Call Rs1.MoveNext
      ElseIf II.PART_NO > EI.PART_NO Then
         Set O = EI
         Call Rs2.MoveNext
      ElseIf II.PART_NO = EI.PART_NO Then
         Set O = II
         Call Rs1.MoveNext
         Call Rs2.MoveNext
      End If
   End If
      
   Set Pi = New CPartItem
   Pi.PART_ITEM_ID = O.PART_ITEM_ID
   Pi.PART_NO = O.PART_NO
   Pi.PART_DESC = O.PART_DESC
   Pi.PART_TYPE_NO = O.PART_TYPE_NO
   Pi.PART_TYPE_NAME = O.PART_TYPE_NAME
   Set GetMinPartItem1 = Pi
   
   Set II = Nothing
   Set EI = Nothing
End Function

Public Sub LoadInventoryBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PigFlag As String = "N")
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim E As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim TempExp As CExportItem
Dim NewDate As Date

   Set D = New CImportItem
   Set E = New CExportItem
   Set Rs = New ADODB.Recordset

   NewDate = DateAdd("D", -1, FromDate)
   
   D.FROM_DATE = -1
   D.TO_DATE = InternalDateToDate(DateToStringIntHi(NewDate))
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = PigFlag
   D.OrderBy = 1
   Call D.QueryData(8, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)) & "-" & Trim(Str(TempData.LOCATION_ID)) & "-" & Trim(Str(TempData.TRANSACTION_SEQ)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   '====
   E.TO_DATE = InternalDateToDate(DateToStringIntHi(NewDate))
   E.COMMIT_FLAG = CommitFlag
   E.LOCATION_ID = LocationID
   E.PIG_FLAG = PigFlag
   E.OrderBy = 1
   Call E.QueryData(18, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempExp = New CExportItem
      Call TempExp.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempExp, Trim(Str(TempExp.PART_ITEM_ID)) & "-" & Trim(Str(TempExp.LOCATION_ID)) & "-" & Trim(Str(TempExp.TRANSACTION_SEQ)))
      End If
      
      Set TempExp = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInventoryBalanceEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartItemID As Long = -1, Optional BatchID As Long = -1, Optional PigFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim NewDate As Date

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   NewDate = DateAdd("D", -1, FromDate)

   D.FROM_DATE = -1
   D.TO_DATE1 = InternalDateToDate(DateToStringIntHi(NewDate))
   D.PART_ITEM_ID = PartItemID
'   D.PART_ITEM_ID = 14469
   D.LOCATION_ID = LocationID
   D.BATCH_ID = BatchID
   D.PIG_FLAG = PigFlag
   D.OrderBy = 1
   'Call D.QueryData(6, Rs, ItemCount) 'เปลี่ยนจาก ind 6 เป็น 21
   Call D.QueryData(21, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      'Call TempData.PopulateFromRS(6, Rs) 'เปลี่ยนจาก ind 6 เป็น 21
      Call TempData.PopulateFromRS(21, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         '''debug.print (TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.BALANCE_AMOUNT)
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInventoryBalanceExByPart(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartItemID As Long = -1, Optional BatchID As Long = -1, Optional PigFlag As String = "", Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim NewDate As Date

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   NewDate = DateAdd("D", -1, FromDate)

   D.FROM_DATE = -1
   D.TO_DATE1 = InternalDateToDate(DateToStringIntHi(NewDate))
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.BATCH_ID = BatchID
   D.PIG_FLAG = PigFlag
   D.PART_TYPE = PartType
   D.OrderBy = 1
   'Call D.QueryData(13, Rs, ItemCount)  'เปลี่ยนจาก ind 13 เป็น 22
   Call D.QueryData(22, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      'Call TempData.PopulateFromRS(13, Rs)
      Call TempData.PopulateFromRS(22, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadInventoryBalanceForBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartItemID As Long = -1, Optional BatchID As Long = -1, Optional PigFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim NewDate As Date

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   NewDate = DateAdd("D", -1, FromDate)

   D.FROM_DATE = -1
   D.TO_DATE1 = InternalDateToDate(DateToStringIntHi(NewDate))
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.BATCH_ID = BatchID
   D.PIG_FLAG = PigFlag
   D.OrderBy = 1
   'Call D.QueryData(14, Rs, ItemCount) 'เปลี่ยนจาก ind 14 เป็น 23
   Call D.QueryData(23, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      'Call TempData.PopulateFromRS(14, Rs)
      Call TempData.PopulateFromRS(23, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   If Rs.State = adStateOpen Then
      Call Rs.Close
   End If
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumPigBalanceAmountYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional BatchID As Long = -1, Optional PigType As String = "")
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE1 = ToDate
   D.PART_ITEM_ID = -1
   D.LOCATION_ID = LocationID
   D.BATCH_ID = BatchID
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   D.PIG_TYPE = PigType
   Call D.QueryData(9, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM & "-" & TempData.LOCATION_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumLocationBalanceAmountYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE1 = ToDate
   D.PART_ITEM_ID = -1
   D.LOCATION_ID = LocationID
   D.BATCH_ID = BatchID
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   Call D.QueryData(10, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExportInventoryBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PigFlag As String = "N")
On Error GoTo ErrorHandler
Dim E As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim TempExp As CExportItem
Dim NewDate As Date

   Set E = New CExportItem
   Set Rs = New ADODB.Recordset
   
   E.FROM_DATE = FromDate
   E.TO_DATE = ToDate
   E.COMMIT_FLAG = CommitFlag
   E.LOCATION_ID = LocationID
   E.PIG_FLAG = PigFlag
   E.OrderBy = 1
   Call E.QueryData(18, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempExp = New CExportItem
      Call TempExp.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempExp, DateToStringInt(TempExp.DOCUMENT_DATE) & "-" & TempExp.LOCATION_ID & "-" & TempExp.PART_ITEM_ID)
      End If
      
      Set TempExp = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportInventoryBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PigFlag As String = "N")
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim E As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim NewDate As Date

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = PigFlag
   D.OrderBy = 1
   Call D.QueryData(8, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, DateToStringInt(TempData.DOCUMENT_DATE) & "-" & TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplier(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSupplier
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CSupplier
Dim I As Long

   Set D = New CSupplier
   D.SUPPLIER_ID = -1
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData2(Rs, itemcount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplier
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_NAME)
         C.ItemData(I) = TempData.SUPPLIER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.SUPPLIER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   'Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CAccount
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CAccount
Dim I As Long

   Set D = New CAccount
   D.ACCOUNT_ID = -1
'   D.CUSTOMER_ID = CustomerID
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(1, Rs, itemcount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CAccount
      Call TempData.PopulateFromRS(1, Rs)
      If ((TempData.CUSTOMER_ID = CustomerID) And (TempData.CUSTOMER_ID > 0)) Or (CustomerID <= 0) Then
         If Not (C Is Nothing) Then
            I = I + 1
            C.AddItem (TempData.ACCOUNT_NO)
            C.ItemData(I) = TempData.ACCOUNT_ID
         End If
      
         If Not (Cl Is Nothing) Then
            Call Cl.Add(TempData, Trim(Str(TempData.ACCOUNT_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
'   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAccountEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CAccount
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CAccount
Dim I As Long

   Set D = New CAccount
   D.ACCOUNT_ID = -1
   D.CUSTOMER_ID = CustomerID
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(1, Rs, itemcount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCOUNT_NO)
         C.ItemData(I) = TempData.ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.ACCOUNT_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBillingDocDistinctAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerCode As String = "")
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set D = New CBillingDoc
   D.BILLING_DOC_ID = -1
   D.CUSTOMER_CODE = CustomerCode
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(2, Rs, itemcount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCOUNT_NO)
         C.ItemData(I) = TempData.ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.ACCOUNT_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBillingDocCode(D As CBillingDoc, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   D.BILLING_DOC_ID = -1

   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(1, Rs, itemcount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCOUNT_NO)
         C.ItemData(I) = TempData.ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.DOCUMENT_NO & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomerAddress(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerID As Long = -1, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Static D As CAddress
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long
   
   If Rs Is Nothing Then
      Set D = New CAddress
      Set Rs = New ADODB.Recordset
      
      D.ENTERPRISE_ID = -1
      D.CUSTOMER_ID = CustomerID
      Call D.QueryData3(Rs, itemcount)
   Else
      If (D.CUSTOMER_ID <> CustomerID) Then
         D.ENTERPRISE_ID = -1
         D.CUSTOMER_ID = CustomerID
         Call D.QueryData3(Rs, itemcount)
      End If
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(Rs)
      
      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.ADDRESS_ID
      End If
      If (I > 0) And ShowFirst Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Not Rs.BOF Then
      Rs.MoveFirst
   End If
   
'   Set Rs = Nothing
'   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplierAddress(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierID As Long = -1, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long

   Set D = New CAddress
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = -1
   D.SUPPLIER_ID = SupplierID
   Call D.QueryData4(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.ADDRESS_ID
      End If
      If (I > 0) And ShowFirst Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadEnterpriseAddress(C As ComboBox, Optional Cl As Collection = Nothing, Optional EnterpriseID As Long = -1, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long
Dim TempIndex As Long

   TempIndex = 0
   
   If Rs Is Nothing Then
      Set D = New CAddress
      Set Rs = New ADODB.Recordset
   
      D.ENTERPRISE_ID = EnterpriseID
      Call D.QueryData2(Rs, itemcount)
   End If
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.ADDRESS_ID
      End If
      If (I > 0) And ShowFirst Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Rs.MoveFirst
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomer(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCustomer
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CCustomer
Dim I As Long

   Set D = New CCustomer
   D.CUSTOMER_ID = -1
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData2(Rs, itemcount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomer
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CUSTOMER_NAME)
         C.ItemData(I) = TempData.CUSTOMER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   'Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadEmployee(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CEmployee
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CEmployee
Dim I As Long

   If Rs Is Nothing Then
      Set D = New CEmployee
      Set Rs = New ADODB.Recordset
      
      D.EMP_ID = -1
      Call D.QueryData(Rs, itemcount)
   Else
      Rs.MoveFirst
   End If
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEmployee
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.Name & " " & TempData.LASTNAME)
         C.ItemData(I) = TempData.EMP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.EMP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
'   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadEmployeeCode(C As ComboBox, Optional Cl As Collection = Nothing)
On Error Resume Next
Dim D As CEmployee
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CEmployee
Dim I As Long

   If Rs Is Nothing Then
      Set D = New CEmployee
      Set Rs = New ADODB.Recordset
      
      D.EMP_ID = -1
      Call D.QueryData(Rs, itemcount)
   Else
      Rs.MoveFirst
   End If
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEmployee
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.Name & " " & TempData.LASTNAME)
         C.ItemData(I) = TempData.EMP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.EMP_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
'   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadProductStatus(C As ComboBox, Optional Cl As Collection = Nothing, Optional PigStatus As Long = -1)
On Error GoTo ErrorHandler
Dim D As CProductStatus
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CProductStatus
Dim I As Long
   
   If Rs Is Nothing Then
      Set D = New CProductStatus
      Set Rs = New ADODB.Recordset
   
      D.PRODUCT_STATUS_ID = PigStatus
      Call D.QueryData(Rs, itemcount)
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CProductStatus
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PRODUCT_STATUS_NAME)
         C.ItemData(I) = TempData.PRODUCT_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PRODUCT_STATUS_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Rs.MoveFirst
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadProductStatusCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional PigStatus As Long = -1)
On Error GoTo ErrorHandler
Dim D As CProductStatus
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CProductStatus
Dim I As Long
   
   If Rs Is Nothing Then
      Set D = New CProductStatus
      Set Rs = New ADODB.Recordset
   
      D.PRODUCT_STATUS_ID = PigStatus
      Call D.QueryData(Rs, itemcount)
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CProductStatus
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PRODUCT_STATUS_NAME)
         C.ItemData(I) = TempData.PRODUCT_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PRODUCT_STATUS_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Rs.MoveFirst
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadProductType(C As ComboBox, Optional Cl As Collection = Nothing, Optional ParentFlag As String = "")
On Error GoTo ErrorHandler
Static D As CProductType
Dim itemcount As Long
Static Rs_ProDuctType As ADODB.Recordset
Dim TempData As CProductType
Dim I As Long

   
   If Rs_ProDuctType Is Nothing Then
      Set D = New CProductType
      Set Rs_ProDuctType = New ADODB.Recordset
   
      D.PRODUCT_TYPE_ID = -1
      D.CAPITAL_FLAG = ParentFlag
      D.OrderType = 1
      Call D.QueryData(Rs_ProDuctType, itemcount)
   Else
      If D.CAPITAL_FLAG <> ParentFlag Then
         D.PRODUCT_TYPE_ID = -1
         D.CAPITAL_FLAG = ParentFlag
         D.OrderType = 1
         Call D.QueryData(Rs_ProDuctType, itemcount)
      End If
   End If
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs_ProDuctType.EOF
      I = I + 1
      Set TempData = New CProductType
      Call TempData.PopulateFromRS(1, Rs_ProDuctType)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PRODUCT_TYPE_NAME & " (" & TempData.PRODUCT_TYPE_NO & ")")
         C.ItemData(I) = TempData.PRODUCT_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PRODUCT_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs_ProDuctType.MoveNext
   Wend
   
   If Not Rs_ProDuctType.BOF Then
      Rs_ProDuctType.MoveFirst
   End If

   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadProductTypeEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional ParentFlag As String = "")
On Error GoTo ErrorHandler
Static D As CProductType
Dim itemcount As Long
Static Rs_ProDuctType As ADODB.Recordset
Dim TempData As CProductType
Dim I As Long

   
   If Rs_ProDuctType Is Nothing Then
      Set D = New CProductType
      Set Rs_ProDuctType = New ADODB.Recordset
   
      D.PRODUCT_TYPE_ID = -1
      D.CAPITAL_FLAG = ParentFlag
      Call D.QueryData(Rs_ProDuctType, itemcount)
   Else
      If D.CAPITAL_FLAG <> ParentFlag Then
         D.PRODUCT_TYPE_ID = -1
         D.CAPITAL_FLAG = ParentFlag
         Call D.QueryData(Rs_ProDuctType, itemcount)
      End If
   End If
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs_ProDuctType.EOF
      I = I + 1
      Set TempData = New CProductType
      Call TempData.PopulateFromRS(1, Rs_ProDuctType)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PRODUCT_TYPE_NAME & " (" & TempData.PRODUCT_TYPE_NO & ")")
         C.ItemData(I) = TempData.PRODUCT_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PRODUCT_TYPE_NO))
      End If
      
      Set TempData = Nothing
      Rs_ProDuctType.MoveNext
   Wend
   
   If Not Rs_ProDuctType.BOF Then
      Rs_ProDuctType.MoveFirst
   End If

   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadProductStatusEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional StatusGroupID As Long = -1, Optional FlagType As String)
On Error GoTo ErrorHandler
Dim D As CStatusGroup
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSGroupItem
Dim Sgi As CSGroupItem
Dim I As Long
Dim IsOK As Boolean

   Set D = New CStatusGroup
   Set Rs = New ADODB.Recordset
   
   D.STATUS_GROUP_ID = StatusGroupID
   D.QueryFlag = 1
   Call glbMaster.QueryStatusGroup(D, Rs, itemcount, IsOK, glbErrorLog)
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   For Each Sgi In D.HGroupItems
      I = I + 1
      Set TempData = New CSGroupItem
      Call TempData.CopyField(1, Sgi)
      
      If Not (C Is Nothing) Then
         C.AddItem (Sgi.STATUS_NAME)
         C.ItemData(I) = Sgi.ST_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If Len(FlagType) > 0 Then
            If TempData.SELECT_FLAG = FlagType Then
               Call Cl.Add(TempData, Trim(Str(TempData.ST_STATUS_ID)))
            End If
         Else
            Call Cl.Add(TempData, Trim(Str(TempData.ST_STATUS_ID)))
         End If
      End If
      
      Set TempData = Nothing
   Next Sgi
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function PigCodeToID_Old(Cd As String) As Long
   If Cd = "N" Then
      PigCodeToID_Old = 1
   ElseIf Cd = "B" Then
      PigCodeToID_Old = 2
   ElseIf Cd = "BT" Then
      PigCodeToID_Old = 3
   ElseIf Cd = "G" Then
      PigCodeToID_Old = 4
   ElseIf Cd = "L" Then
      PigCodeToID_Old = 5
   ElseIf Cd = "R" Then
      PigCodeToID_Old = 6
   End If
End Function

Public Function PigCodeToID(Cd As String) As Long
Dim TempCol As Collection
Dim Pt As CProductType

   Set TempCol = New Collection
   Call LoadProductType(Nothing, TempCol)
      
   For Each Pt In TempCol
      If Pt.PRODUCT_TYPE_NO = Cd Then
         PigCodeToID = Pt.PRODUCT_TYPE_ID
         Exit Function
      End If
   Next Pt
   
   PigCodeToID = -1
   Set TempCol = Nothing
End Function

Public Function PigTypeToCode(Cd As Long) As String
Static TempCol As Collection
Static I As Long
Dim Pt As CProductType
   
   If I = 0 Then
      Set TempCol = New Collection
      Call LoadProductType(Nothing, TempCol)
   End If
   
   For Each Pt In TempCol
      If Pt.PRODUCT_TYPE_ID = Cd Then
         PigTypeToCode = Pt.PRODUCT_TYPE_NO
         Exit Function
      End If
   Next Pt
   
   PigTypeToCode = ""
End Function

'Public Sub PopulatePigType(TempID As Long, PigType As CPigType)
'   If TempID = 1 Then
'      PigType.PIG_TYPE_ID = 1
'      PigType.PIG_TYPE_NO = "N"
'      PigType.PIG_TYPE_NAME = "หมูปกติ (N)"
'   ElseIf TempID = 2 Then
'      PigType.PIG_TYPE_ID = 2
'      PigType.PIG_TYPE_NO = "B"
'      PigType.PIG_TYPE_NAME = "หมูพ่อพันธ์ (B)"
'   ElseIf TempID = 3 Then
'      PigType.PIG_TYPE_ID = 3
'      PigType.PIG_TYPE_NO = "BT"
'      PigType.PIG_TYPE_NAME = "หมูสำรองพ่อ (BT)"
'   ElseIf TempID = 4 Then
'      PigType.PIG_TYPE_ID = 4
'      PigType.PIG_TYPE_NO = "G"
'      PigType.PIG_TYPE_NAME = "แม่อุ้มท้อง (G)"
'   ElseIf TempID = 5 Then
'      PigType.PIG_TYPE_ID = 5
'      PigType.PIG_TYPE_NO = "L"
'      PigType.PIG_TYPE_NAME = "แม่คลอด (L)"
'   ElseIf TempID = 6 Then
'      PigType.PIG_TYPE_ID = 6
'      PigType.PIG_TYPE_NO = "R"
'      PigType.PIG_TYPE_NAME = "สำรองแม่ (R)"
'   End If
'
'   PigType.KEY_ID = PigType.PIG_TYPE_ID
'   PigType.KEY_LOOKUP = PigType.PIG_TYPE_NO
'End Sub

'==
'Public Sub LoadProductType(C As ComboBox, Optional Cl As Collection = Nothing)
'On Error GoTo ErrorHandler
'Dim ItemCount As Long
'Dim I As Long
'Dim TempData As CPigType
'
'   If Not (C Is Nothing) Then
'      C.Clear
'      I = 0
'      C.AddItem ("")
'   End If
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
'
'   For I = 1 To 6
'      Set TempData = New CPigType
'      Call PopulatePigType(I, TempData)
'
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.PIG_TYPE_NAME)
'         C.ItemData(I) = TempData.PIG_TYPE_ID
'      End If
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.Add(TempData, Trim(Str(TempData.PIG_TYPE_ID)))
'      End If
'
'      Set TempData = Nothing
'   Next I
'
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub


Public Function DocumentTypeToString(Dt As Long) As String
   If Dt = 1 Then
      DocumentTypeToString = "ใบนำเข้า"
   ElseIf Dt = 2 Then
      DocumentTypeToString = "ใบเบิกวัตถุดิบ"
   ElseIf Dt = 3 Then
      DocumentTypeToString = "ใบโอนวัตถุดิบ"
   ElseIf Dt = 4 Then
      DocumentTypeToString = "ใบปรับยอด"
   End If
End Function

Public Function CommitTypeToFlag(Ct As Long) As String
   If Ct = 1 Then
      CommitTypeToFlag = "Y"
   ElseIf Ct = 2 Then
      CommitTypeToFlag = "N"
   Else
      CommitTypeToFlag = ""
   End If
End Function

Public Function ID2Orientation(TempID As OrientationSettings) As String
   If TempID = orLandscape Then
      ID2Orientation = "แนวนอน"
   Else
      ID2Orientation = "แนวตั้ง"
   End If
End Function

Public Function ID2PaperSize(TempID As PaperSizeSettings) As String
   If TempID = pprA4 Then
      ID2PaperSize = "A4"
   ElseIf TempID = pprLetter Then
      ID2PaperSize = "Letter"
   ElseIf TempID = pprFanfoldUS Then
      ID2PaperSize = "Us standard"
   Else
      ID2PaperSize = "A4"
   End If
End Function

Public Sub InitPaperSize(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2PaperSize(pprA4))
   C.ItemData(1) = pprA4

   C.AddItem (ID2PaperSize(pprLetter))
   C.ItemData(2) = pprLetter

   C.AddItem (ID2PaperSize(pprFanfoldUS))
   C.ItemData(3) = pprFanfoldUS
End Sub

Public Sub InitFontName(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("AngsanaUPC")
   C.ItemData(1) = 1
End Sub

Public Sub InitOrientation(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2Orientation(orLandscape))
   C.ItemData(1) = orLandscape

   C.AddItem (ID2Orientation(orPortrait))
   C.ItemData(2) = orPortrait
End Sub

Public Sub InitLossType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("% สูญเสียต่อวัน")
   C.ItemData(1) = 1

   C.AddItem ("จำนวนตัวที่สูญเสีย")
   C.ItemData(2) = 2

   C.AddItem ("จำนวนตัวที่เหลือจากสูญเสีย")
   C.ItemData(3) = 3
End Sub
Public Function GetLossType(ID As Long) As String
   If ID = 1 Then
      GetLossType = "% สูญเสียต่อวัน"
   ElseIf ID = 2 Then
      GetLossType = "จำนวนตัวที่สูญเสีย"
   ElseIf ID = 3 Then
      GetLossType = "จำนวนตัวที่เหลือจากสูญเสีย"
   End If
End Function

Public Sub InitSellShareType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("แบ่งตาม %")
   C.ItemData(1) = 1

   C.AddItem ("กำหนดจำนวนตัว")
   C.ItemData(2) = 2

   C.AddItem ("เอาทั้งหมดที่มี")
   C.ItemData(3) = 3
End Sub
Public Function GetSellShareType(ID As Long) As String
   If ID = 1 Then
      GetSellShareType = "แบ่งตาม %"
   ElseIf ID = 2 Then
      GetSellShareType = "กำหนดจำนวนตัว"
   ElseIf ID = 3 Then
      GetSellShareType = "เอาทั้งหมดที่มี"
   End If
End Function
Public Sub LoadAgeRange(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CAgeRange
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAgeRange
Dim I As Long

   Set D = New CAgeRange
   Set Rs = New ADODB.Recordset
   
   D.AGE_RANGE_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAgeRange
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.AGE_RANGE_NAME)
         C.ItemData(I) = TempData.AGE_RANGE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.AGE_RANGE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAccessRight(C As ComboBox, Optional Cl As Collection = Nothing, Optional GroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CGroupRight
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGroupRight
Dim I As Long

   Set D = New CGroupRight
   Set Rs = New ADODB.Recordset
   
   D.GROUP_RIGHT_ID = -1
   D.GROUP_ID = GroupID
   Call D.QueryData3(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CGroupRight
      Call TempData.PopulateFromRS3(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.RIGHT_ITEM_NAME)
         C.ItemData(I) = TempData.GROUP_RIGHT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRelatedImportItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional FromGuiID As Long, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.FROM_GUI_ID = FromGuiID
   D.TO_GUI_ID = -1
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GUI_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRelatedExportItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional FromGuiID As Long, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.FROM_GUI_ID = FromGuiID
   D.TO_GUI_ID = -1
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GUI_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function PaymentTypeToText(ID As PAYMENT_TYPE) As String
   If ID = CASH_PMT Then
      PaymentTypeToText = MapText("เงินสด")
   ElseIf ID = CHECK_PMT Then
      PaymentTypeToText = MapText("เช็ค")
   ElseIf ID = CREDITCRD_PMT Then
      PaymentTypeToText = MapText("บัตรเครดิต")
   ElseIf ID = BANKTRF_PMT Then
      PaymentTypeToText = MapText("โอนผ่านธนาคาร")
   ElseIf ID = CASHRET_PMT Then
      PaymentTypeToText = MapText("เงินสดขายคืน")
   End If
End Function

Public Sub InitPaymentType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (PaymentTypeToText(CASH_PMT))
   C.ItemData(1) = CASH_PMT
   
   C.AddItem (PaymentTypeToText(CHECK_PMT))
   C.ItemData(2) = CHECK_PMT

   C.AddItem (PaymentTypeToText(BANKTRF_PMT))
   C.ItemData(3) = BANKTRF_PMT
End Sub

Public Sub InitPaymentTypeEx(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (PaymentTypeToText(CASH_PMT))
   C.ItemData(1) = CASH_PMT
   
   C.AddItem (PaymentTypeToText(CHECK_PMT))
   C.ItemData(2) = CHECK_PMT

   C.AddItem (PaymentTypeToText(BANKTRF_PMT))
   C.ItemData(3) = BANKTRF_PMT
   
   C.AddItem (PaymentTypeToText(CASHRET_PMT))
   C.ItemData(4) = CASHRET_PMT
End Sub

Public Sub LoadBank(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CBank
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBank
Dim I As Long

   Set D = New CBank
   Set Rs = New ADODB.Recordset
   
   D.BANK_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBank
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.BANK_NAME)
         C.ItemData(I) = TempData.BANK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.BANK_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBankAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional BankID As Long = -1, Optional BankBranch As Long = -1, Optional KeyType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBankAccount
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBankAccount
Dim I As Long

   Set D = New CBankAccount
   Set Rs = New ADODB.Recordset
   
   D.BANK_ACCOUNT_ID = -1
   D.BANK_ID = BankID
   D.BBRANCH_ID = BankBranch
   
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBankAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCOUNT_NAME)
         C.ItemData(I) = TempData.BANK_ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.Add(TempData, Trim(TempData.ACCOUNT_NAME))
         Else
            Call Cl.Add(TempData, Trim(Str(TempData.BANK_ACCOUNT_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadBankBranch(C As ComboBox, Optional Cl As Collection = Nothing, Optional BankID As Long)
On Error GoTo ErrorHandler
Dim D As CBankBranch
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBankBranch
Dim I As Long

   Set D = New CBankBranch
   Set Rs = New ADODB.Recordset
   
   D.BBRANCH_ID = -1
   D.BANK_ID = BankID
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBankBranch
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.BBRANCH_NAME)
         C.ItemData(I) = TempData.BBRANCH_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.BBRANCH_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBankBranchEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional BankID As Long)
On Error GoTo ErrorHandler
Dim D As CBankBranch
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBankBranch
Dim I As Long

   Set D = New CBankBranch
   Set Rs = New ADODB.Recordset
   
   D.BBRANCH_ID = -1
   D.BANK_ID = BankID
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBankBranch
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.BBRANCH_NAME)
         C.ItemData(I) = TempData.BBRANCH_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.BBRANCH_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRegion(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CRegion
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRegion
Dim I As Long

   Set D = New CRegion
   Set Rs = New ADODB.Recordset
   
   D.REGION_ID = -1
   D.REGION_NAME = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CRegion
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.REGION_NAME)
         C.ItemData(I) = TempData.REGION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.REGION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadRegionEx(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CRegion
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRegion
Dim I As Long

   Set D = New CRegion
   Set Rs = New ADODB.Recordset
   
   D.REGION_ID = -1
   D.REGION_NAME = ""
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CRegion
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.REGION_NAME)
         C.ItemData(I) = TempData.REGION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.REGION_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadReceiptByDocTypeSubTypeReceiptType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadReceiptByDocTypeSubTypeReceiptTypeAcc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional ValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.VALID_DATE = ValidDate
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.ACCOUNT_ID & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSellPriceByDocTypeSubTypeReceiptType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSellRevenuePriceByDocTypeSubTypeReceiptType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(8, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
'''debug.print TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE & ":" & TempData.TOTAL_PRICE
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSellPriceByDocTypeSubTypeReceiptTypeAcc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional ValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.VALID_DATE = ValidDate
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.ACCOUNT_ID & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaidAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaidAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1, Optional ValidDate As Date = -1, Optional CustomerCode As String = "")
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.VALID_DATE = ValidDate
   D.CUSTOMER_CODE = CustomerCode
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadAllPaidAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1, Optional ValidDate As Date = -1, Optional CustomerCode As String = "")
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim Bd As CBillingDoc
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.VALID_DATE = ValidDate
   D.CUSTOMER_CODE = CustomerCode
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(17, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(17, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set Bd = GetObject("CBillingDoc", Cl, Trim(Str(TempData.DO_ID)), False)
         If Bd Is Nothing Then
            Set Bd = New CBillingDoc
            Bd.BILLING_DOC_ID = TempData.DO_ID
            Call Cl.Add(Bd, Trim(Str(Bd.BILLING_DOC_ID)))
         Else
            ''debug.print
         End If
         Call Bd.ReceiptItems.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDnCnAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional SourceType As Long = 1, Optional ValidDate As Date = -1, Optional CustomerCode As String = "")
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   If SourceType = 1 Then
      D.FROM_ITEM_DATE = FromDate
      D.TO_ITEM_DATE = ToDate
   ElseIf SourceType = 2 Then
      D.FROM_DOC_DATE = FromDate
      D.TO_DOC_DATE = ToDate
   End If
   D.DOCUMENT_TYPE = DocumentType
   D.VALID_DATE = ValidDate
   D.CUSTOMER_CODE = CustomerCode
   Call D.QueryData(7, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.DO_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDnCnAmountByAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional DateType As Long = 1, Optional ValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   If DateType = 1 Then
      D.FROM_ITEM_DATE = FromDate
      D.TO_ITEM_DATE = ToDate
   ElseIf DateType = 2 Then
      D.FROM_DOC_DATE = FromDate
      D.TO_DOC_DATE = ToDate
   End If
   D.DOCUMENT_TYPE = DocumentType
   D.VALID_DATE = ValidDate
   Call D.QueryData(6, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.ACCOUNT_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDnCnAmountByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional ValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   'D.FROM_ITEM_DATE = FromDate
   'D.TO_ITEM_DATE = ToDate
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   D.VALID_DATE = ValidDate
   Call D.QueryData(9, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDnCnAmountByDocType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional ValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   D.VALID_DATE = ValidDate
   Call D.QueryData(8, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.DOCUMENT_TYPE)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaidAmountByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1, Optional ValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.VALID_DATE = ValidDate
   Call D.QueryData(5, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalPriceByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RcpType As Long = -1, Optional ValidDate As Date = -1, Optional CustomerCode As String = "")
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.RECEIPT_TYPE = RcpType
   D.VALID_DATE = ValidDate
   D.CUSTOMER_CODE = CustomerCode
   Call D.QueryData(5, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctSellGroupType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RcpType As Long = -1, Optional ValidDate As Date = -1, Optional CustomerCode As String = "")
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.RECEIPT_TYPE = RcpType
   D.VALID_DATE = ValidDate
   D.CUSTOMER_CODE = CustomerCode
   D.DOCUMENT_TYPE = 1
   Call D.QueryData(37, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(37, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CDoItem", Cl, Trim(Str(TempData.DO_ID)), False)
         If D Is Nothing Then
            Call Cl.Add(TempData, Trim(Str(TempData.DO_ID)))
         Else
            D.PART_GROUP_NAME = D.PART_GROUP_NAME & "/" & TempData.PART_GROUP_NAME
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalROPrice(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CROItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CROItem
Dim I As Long

   Set D = New CROItem
   Set Rs = New ADODB.Recordset
   
   D.RO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   Call D.QueryData(6, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CROItem
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalExpenseCash(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.IMPORT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = "N"
   Call D.QueryData(26, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(26, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM2 & "-" & TempData.PART_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigUsedYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.IMPORT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = "N"
   D.BATCH_ID = BatchID
   Call D.QueryData(27, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(27, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigImportAmountYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.IMPORT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   D.BATCH_ID = BatchID
   Call D.QueryData(27, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(27, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.YYYYMM & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumAmountByBillDocTypeSubType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatus As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_STATUS = PigStatus
   Call D.QueryData(13, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(13, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DO_ID & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatus As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_STATUS = PigStatus
   Call D.QueryData(14, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigSellPriceYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1, Optional PartNo As String)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = "Y"
   D.BATCH_ID = BatchID
   D.PART_NO = PartNo
   Call D.QueryData(20, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(20, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_STATUS & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRevenueSellPriceYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = ""
   D.BATCH_ID = BatchID
   Call D.QueryData(21, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(21, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.REVENUE_ID & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalPriceByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional ValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   D.VALID_DATE = ValidDate
   Call D.QueryData(6, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDueDateInterval(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -30
   MM.MAX = 0
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = -60
   MM.MAX = -30
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = -60
   Call Cl.Add(MM)
   Set MM = Nothing
   
   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 15
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 15
   MM.MAX = 30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 30
   MM.MAX = 60
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 60
   MM.MAX = 90
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = 90
   MM.MAX = 9999999
   Call Cl.Add(MM)
   Set MM = Nothing
End Sub

Public Sub LoadPackageType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPackageType
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPackageType
Dim I As Long

   Set D = New CPackageType
   Set Rs = New ADODB.Recordset
   
   D.PACKAGE_TYPE_ID = -1
   
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPackageType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PACKAGE_TYPE_NAME)
         C.ItemData(I) = TempData.PACKAGE_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PACKAGE_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitPackageOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสการตั้งราคา")
   C.ItemData(1) = 1
   
   C.AddItem ("รายละเอียด")
   C.ItemData(2) = 2
   
   C.AddItem ("ประเภทการตั้งราคา")
   C.ItemData(3) = 3
   
End Sub

Public Sub LoadPackage(C As ComboBox, Optional Cl As Collection = Nothing, Optional PackageType As Long = -1, Optional BasicFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CPackage
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPackage
Dim I As Long

   Set D = New CPackage
   Set Rs = New ADODB.Recordset
   
   D.PKG_ID = -1
   D.PKG_TYPE = PackageType
   D.PKG_BASIC_FLAG = BasicFlag
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPackage
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PKG_NAME)
         C.ItemData(I) = TempData.PKG_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PKG_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Public Sub LoadCustomerPackage(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerIS As Long = -1, Optional PackageType As Long = -1)
'On Error GoTo ErrorHandler
'Dim D As CCustomerPackage
'Dim Itemcount As Long
'Dim Rs As ADODB.Recordset
'Dim TempData As CCustomerPackage
'Dim I As Long
'
'   Set D = New CCustomerPackage
'   Set Rs = New ADODB.Recordset
'
'   D.CUSTOMER_PACKAGE_ID = -1
'   D.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'   D.PKG_TYPE = uctlPackageType.MyCombo.ItemData(Minus2Zero(uctlPackageType.MyCombo.ListIndex))
'
'   Call D.QueryData(1, Rs, Itemcount)
'
'   If Not (C Is Nothing) Then
'      C.Clear
'      I = 0
'      C.AddItem ("")
'   End If
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
'   While Not Rs.EOF
'      I = I + 1
'      Set TempData = New CPackage
'      Call TempData.PopulateFromRS(1, Rs)
'
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.PKG_NAME)
'         C.ItemData(I) = TempData.PKG_ID
'      End If
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.Add(TempData, Trim(Str(TempData.PKG_ID)))
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'
'   Set Rs = Nothing
'   Set D = Nothing
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub
Public Sub InitPackageType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem ("คิดแบบมีค่าสายพันธ์ (บวกค่าคงที่)")
   C.ItemData(1) = 4
   
   C.AddItem ("คิดแบบขึ้นกับน้ำหนัก")
   C.ItemData(2) = 5
   
End Sub

Public Function IDToPackageName(ID As Long) As String
   If ID = 1 Then
      IDToPackageName = "หมูพันธ์"
   ElseIf ID = 2 Then
      IDToPackageName = "แบบเหมา"
   ElseIf ID = 3 Then
      IDToPackageName = "คิดแบบขึ้นกับน้ำหนัก"
   ElseIf ID = 4 Then
      IDToPackageName = "คิดแบบมีค่าสายพันธ์ (บวกค่าคงที่)"
   ElseIf ID = 5 Then
      IDToPackageName = "คิดแบบขึ้นกับน้ำหนัก"
   ElseIf ID = 6 Then
      IDToPackageName = "ค่าสายพันธ์+ประเภทน้ำหนัก"
   End If
End Function
Public Sub LoadCustomerPackage(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCustomerPackage
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomerPackage
Dim I As Long

   Set D = New CCustomerPackage
   Set Rs = New ADODB.Recordset
   
   D.CUSTOMER_PACKAGE_ID = -1
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomerPackage
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PKG_NAME)
         C.ItemData(I) = TempData.CUSTOMER_PACKAGE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSTOMER_ID & "-" & TempData.PKG_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPackageDetail(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPackageDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPackageDetail
Dim I As Long

   Set D = New CPackageDetail
   Set Rs = New ADODB.Recordset
   
   D.PKG_DETAIL_ID = -1
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPackageDetail
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PKG_DETAIL_ID)
         C.ItemData(I) = TempData.PKG_DETAIL_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadInitialPigBalanceEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim TempData As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim I As Long

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset
   
   D.BALANCE_ACCUM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.LOCATION_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadInitialCostAccum(C As ComboBox, Optional Cl As Collection = Nothing, Optional CL2 As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMovementItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMovementItem
Dim I As Long

Dim TempColl As Collection
Dim Ba As CBalanceAccum
Dim CA As CCost_Accum
Dim CB As CCost_Accum
Dim CostSearch As CCost_Accum

   Set TempColl = New Collection
   Set Ba = New CBalanceAccum
      
   Call LoadInitialPigBalanceEx(Nothing, TempColl, -1, ToDate, , , BatchID)
   
   Set D = New CMovementItem
   Set Rs = New ADODB.Recordset
   
   D.MOVEMENT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_ID = PigID
   D.DOCUMENT_CATEGORY = 3
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   Call D.QueryData(20, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set CL2 = Nothing
      Set Cl = New Collection
      Set CL2 = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMovementItem
      Call TempData.PopulateFromRS(20, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set CA = New CCost_Accum
         CA.LOCATION_ID = TempData.FROM_HOUSE_ID
         CA.PART_ITEM_ID = TempData.PIG_ID
         CA.CUS_ID = -1
         CA.DOCUMENT_CATEGORY = TempData.DOCUMENT_CATEGORY
         CA.DOCUMENT_DATE = TempData.DOCUMENT_DATE
         CA.DOCUMENT_TYPE = TempData.DOCUMENT_TYPE
         CA.BATCH_ID = BatchID
         
         If TempData.PART_ITEM_ID > 0 Then
            CA.COST_RAW = TempData.CAPITAL_AMOUNT
         Else
            CA.COST_EXP = TempData.CAPITAL_AMOUNT
         End If
         
         Set CostSearch = GetCostAccumSearch(Cl, CA.GetKey1)
         If CostSearch Is Nothing Then
            Call Cl.Add(CA, CA.GetKey1)
         Else
            CostSearch.COST_RAW = CostSearch.COST_RAW + CA.COST_RAW
            CostSearch.COST_EXP = CostSearch.COST_EXP + CA.COST_EXP
         End If
         
         Set CB = New CCost_Accum
         
         CB.LOCATION_ID = TempData.FROM_HOUSE_ID
         CB.PART_ITEM_ID = TempData.PIG_ID
         
         Set Ba = GetBalanceAccum(TempColl, Trim(TempData.PIG_ID & "-" & TempData.FROM_HOUSE_ID))
         
         CB.ITEM_AMOUNT = Ba.BALANCE_AMOUNT
         
         If TempData.PART_ITEM_ID > 0 Then
            CB.COST_RAW = TempData.CAPITAL_AMOUNT
         Else
            CB.COST_EXP = TempData.CAPITAL_AMOUNT
         End If

         Set CostSearch = GetCostAccumSearch(CL2, CB.GetKey2)
         If CostSearch Is Nothing Then
            Call CL2.Add(CB, CB.GetKey2)
         Else
            CostSearch.COST_RAW = CostSearch.COST_RAW + CB.COST_RAW
            CostSearch.COST_EXP = CostSearch.COST_EXP + CB.COST_EXP
            'CostSearch.ITEM_AMOUNT = CostSearch.ITEM_AMOUNT + Ba.BALANCE_AMOUNT
         End If
         
         
         Set CA = Nothing
         Set CB = Nothing
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctCostAccumYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCost_Accum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCost_Accum
Dim I As Long

   Set D = New CCost_Accum
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   D.OrderBy = 1
   D.DOCUMENT_TYPE = 10
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCost_Accum
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRelatedImportItemEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional FromGuiID As Long, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.FROM_GUI_ID = FromGuiID
   D.TO_GUI_ID = -1
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(9999, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(9999, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GUI_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRelatedExportItemEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional FromGuiID As Long, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.FROM_GUI_ID = FromGuiID
   D.TO_GUI_ID = -1
   D.OrderBy = 1
   D.BATCH_ID = BatchID
   Call D.QueryData(9999, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(9999, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GUI_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadADGParam(C As ComboBox, Optional Cl As Collection = Nothing, Optional BatchID As Long = -1, Optional ParamArea As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBatchItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBatchItem
Dim I As Long

   Set D = New CBatchItem
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("BATCH_ITEM_ID", -1)
   Call D.SetFieldValue("BATCH_ID", BatchID)
   Call D.SetFieldValue("PARAM_AREA", ParamArea)
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBatchItem
      Call TempData.PopulateFromRS(2, Rs)

      If Not (C Is Nothing) Then
'         C.AddItem (TempData.GetFieldValue("BATCH_NO"))
'         C.ItemData(I) = TempData.GetFieldValue("BATCH_ID")
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSystemParam(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSystemParam
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSystemParam
Dim I As Long

   Set D = New CSystemParam
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("PARAM_ID", -1)
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSystemParam
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.RIGHT_ITEM_NAME)
'         C.ItemData(I) = TempData.GROUP_RIGHT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.GetFieldValue("PARAM_NAME"))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPaymentBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PaymentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPaymentItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPaymentItem
Dim I As Long

   Set D = New CPaymentItem
   Set Rs = New ADODB.Recordset

   D.OrderBy = 1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPaymentItem
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PAYMENT_TYPE & "-" & TempData.TX_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadConfigDoc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CConfigDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CConfigDoc
Dim I As Long

   Set D = New CConfigDoc
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CConfigDoc
      Call TempData.PopulateFromRS(1, Rs)
      
      TempData.Flag = "I"
      
      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.GetFieldValue("CONFIG_DOC_CODE"))
         C.ItemData(I) = TempData.GetFieldValue("CONFIG_DOC_TYPE")
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GetFieldValue("CONFIG_DOC_TYPE"))))
      End If
            
      Set TempData = Nothing
      
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadInventoryBalanceDistinctPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartItemID As Long = -1, Optional BatchID As Long = -1, Optional PigFlag As String = "", Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = -1
   D.TO_DATE = ToDate
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.BATCH_ID = BatchID
   D.PIG_FLAG = PigFlag
   D.PART_TYPE = PartType
   D.OrderBy = 1
   Call D.QueryData(15, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   If Rs.State = adStateOpen Then
      Call Rs.Close
   End If
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function PaymentType2Text(ID As Long) As String
   If ID = 1 Then
      PaymentType2Text = "เงินสด"
   ElseIf ID = 2 Then
      PaymentType2Text = "เงินโอน"
   ElseIf ID = 3 Then
      PaymentType2Text = "เช็ค"
   Else
      PaymentType2Text = ""
   End If
End Function
Public Sub LoadMaster(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterType As MASTER_TYPE, Optional TempID1 As Long = -1, Optional TempID2 As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMasterRef
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long

   Set D = New CMasterRef
   Set Rs = New ADODB.Recordset
   
   D.KEY_ID = -1
   D.MASTER_AREA = MasterType
   D.TEMP_ID1 = TempID1
   D.TEMP_ID2 = TempID2
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.KEY_NAME)
         C.ItemData(I) = TempData.KEY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.KEY_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumCashTrnAmount(Ct As CCashTran, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Ct.SetFieldValue("CASH_TRAN_ID", -1)
   Call Ct.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.GetFieldValue("KEY_NAME"))
'         C.ItemData(I) = TempData.GetFieldValue("KEY_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.GetFieldValue("BANK_ACCOUNT") & "-" & TempData.GetFieldValue("TX_TYPE"))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBankAccountInCashTrn(Ct As CCashTran, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Ct.SetFieldValue("CASH_TRAN_ID", -1)
   Call Ct.SetFieldValue("ORDER_TYPE", 1)
   Call Ct.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.GetFieldValue("KEY_NAME"))
'         C.ItemData(I) = TempData.GetFieldValue("KEY_ID")
      End If
      
      If Not (Cl Is Nothing) Then
'''debug.print TempData.GetFieldValue("BANK_ACCOUNT") & "-" & TempData.GetFieldValue("BANK_ID") & "-" & TempData.GetFieldValue("BANK_BRANCH")
         Call Cl.Add(TempData, TempData.GetFieldValue("BANK_ACCOUNT") & "-" & TempData.GetFieldValue("BANK_ID") & "-" & TempData.GetFieldValue("BANK_BRANCH"))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitReportCashTx(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
End Sub
Public Sub LoadReceiptByBillingDocID(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(12, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.BILLING_DOC_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumCashTranAmountByBillingDocID(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(9, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(9, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GetFieldValue("BILLING_DOC_ID"))))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitReport6_33Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("วันที่รับชำระ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เลขที่ใบเสร็จ"))
   C.ItemData(2) = 2
End Sub
Public Sub InitPaymentType3(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (PaymentTypeToText(CASH_PMT))
   C.ItemData(1) = CASH_PMT
End Sub
Public Sub LoadRemainMoneyBl(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.QueryData(10, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(10, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.GetFieldValue("PAYMENT_TYPE") & "-" & TempData.GetFieldValue("RECEIPT_TYPE")))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadRemainMoneyCd(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.QueryData(11, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(11, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GetFieldValue("PAYMENT_TYPE"))))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBillingDocPayment(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim itemcount As Long
Static Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
Dim D As CBillingDoc
   
   
   Set Rs = New ADODB.Recordset
   Set D = New CBillingDoc
   
   D.BILLING_DOC_ID = -1
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCOUNT_NO)
         C.ItemData(I) = TempData.ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PAYMENT_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctAccountInCashTran(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(13, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(13, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GetFieldValue("BANK_ACCOUNT"))))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumCashTranAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(5, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.GetFieldValue("CUSTOMER_ID") & "-" & TempData.GetFieldValue("BANK_ACCOUNT") & "-" & DateToStringInt(TempData.GetFieldValue("TX_DATE")))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadReceiptByCustomerDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(11, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, DateToStringInt(TempData.DOCUMENT_DATE) & "-" & TempData.CUSTOMER_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumCashTranAmountByCustDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(7, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.GetFieldValue("CUSTOMER_ID") & "-" & TempData.GetFieldValue("PAYMENT_TYPE") & "-" & DateToStringInt(TempData.GetFieldValue("TX_DATE")))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumCashTranAmountByCustDate2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(8, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.GetFieldValue("CUSTOMER_ID") & "-" & TempData.GetFieldValue("PAYMENT_TYPE") & "-" & DateToStringInt(TempData.GetFieldValue("TX_DATE")))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDoIDFromReceiptItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long
Dim PrevID As Long
Dim Seq As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_ITEM_DUE_DATE = FromDate
   D.TO_ITEM_DUE_DATE = ToDate
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(14, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not Rs.EOF Then
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(14, Rs)
      PrevID = TempData.DO_ID
      Seq = 1
      Set TempData = Nothing
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If PrevID <> TempData.DO_ID Then
         Seq = 1
         PrevID = TempData.DO_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DO_ID & "-" & Seq)
'''debug.print TempData.DO_ID & "-" & Seq
         Seq = Seq + 1
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitReport6_41Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่รับชำระ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เลขที่ใบเสร็จ"))
   C.ItemData(2) = 2
End Sub

Public Sub LoadMaxMinReceiptDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long
Dim PrevID As Long
Dim Seq As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_ITEM_DUE_DATE = FromDate
   D.TO_ITEM_DUE_DATE = ToDate
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(15, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
      
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (C Is Nothing) Then
      End If
            
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadChequeFromReceipt(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("PAYMENT_TYPE", 3)
   Call D.SetFieldValue("TX_TYPE", "I")
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(14, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
'''debug.print TempData.GetFieldValue("BILLING_DOC_ID")
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumPigBalanceAmountByAge(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional HouseGroup As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE1 = ToDate
   D.HOUSE_GROUP_ID = HouseGroup
   D.PIG_FLAG = "Y"
   D.OrderBy = 1
   D.LOCATION_ID = LocationID
   Call D.QueryData(16, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(16, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PIG_AGE)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartImportAmountByFeedGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional DocumentType As Long = -1, Optional PartGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.DOCUMENT_TYPE = DocumentType
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "N"
   D.PART_GROUP = PartGroup
   D.OrderBy = 1
   Call D.QueryData(31, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(31, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.FEED_GROUP)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartExportAmountByFeedGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional DocumentType As Long = -1, Optional SaleFlag As String = "", Optional PartGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.PIG_FLAG = "N"
   D.SALE_FLAG = SaleFlag
   D.DOCUMENT_TYPE = DocumentType
   D.PART_GROUP_ID = PartGroup
   D.OrderBy = 1
   Call D.QueryData(58, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(58, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.FEED_GROUP)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartSellAmountByFeedGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.PIG_FLAG = "N"
   D.DocTypeSet = "(10, 13)"
   D.OrderBy = 1
   Call D.QueryData(58, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(58, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.FEED_GROUP)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumBalanceAccum2ByFeedGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = -1
   D.TO_DATE1 = ToDate
   D.OrderBy = 1
   'Call D.QueryData(17, Rs, ItemCount) 'เปลี่ยน 17 เป็น 24
   Call D.QueryData(24, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
     ' Call TempData.PopulateFromRS(17, Rs)
      Call TempData.PopulateFromRS(24, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.FEED_GROUP)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitRevenueCostOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("หมายเลข")
   C.ItemData(1) = 1
   
   C.AddItem ("วันที่")
   C.ItemData(2) = 2
   
End Sub
Public Sub LoadSumPigHouseBalanceLoss(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatusID As Long = -1, Optional LocationID As Long = -1, Optional PigType As String = "")
On Error GoTo ErrorHandler
Dim D As CLossItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLossItem
Dim I As Long

   Set D = New CLossItem
   Set Rs = New ADODB.Recordset
   
   D.LOSS_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_HOUSE_ID = LocationID
   D.PIG_STATUS = PigStatusID
   D.PIG_TYPE = PigType
'      D.DOCUMENT_CATEGORY = 1
'      D.PARENT_FLAG = "N"
'   D.OrderBy = 1
   Call D.QueryData(8, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLossItem
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.FROM_HOUSE_ID & "-" & TempData.PIG_ID & "-" & TempData.PART_GROUP_ID & "-" & TempData.EXPENSE_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMovementLocationLoss(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional YearID As Long = -1, Optional WeekNo As String, Optional PigType As String)
On Error GoTo ErrorHandler
Dim D As CLossItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLossItem
Dim I As Long
Static Locations As Collection
Dim Lc As CLocation

   If Locations Is Nothing Then
      Set Locations = New Collection
      Call LoadLocation(Nothing, Locations, -1, "")
   End If
   
   Set D = New CLossItem
   Set Rs = New ADODB.Recordset
   
   D.LOSS_ITEM_ID = -1
   D.FROM_HOUSE_ID = LocationID
   D.OrderBy = 1
   D.PIG_TYPE = PigType
   D.YEAR_SEQ_ID = YearID
   D.WEEK_NO = WeekNo
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLossItem
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If TempData.FROM_HOUSE_ID > 0 Then
         Set Lc = Locations(Trim(Str(TempData.FROM_HOUSE_ID)))
         
         If Not (Cl Is Nothing) Then
            Call Cl.Add(Lc, Trim(Str(Lc.LOCATION_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Locations = Nothing
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMovementPigLoss(C As ComboBox, Optional Cl As Collection = Nothing, Optional Params As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional ParentFlag As String = "", Optional YearID As Long = -1, Optional WeekNo As String, Optional PigType As String)
On Error GoTo ErrorHandler
Dim D As CLossItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLossItem
Dim I As Long
Static Partitems As Collection
Dim Pi As CPartItem

   If Partitems Is Nothing Then
      Set Partitems = New Collection
      Call LoadPartItem(Nothing, Partitems, , "Y")
   End If

   Set D = New CLossItem
   Set Rs = New ADODB.Recordset
   
   D.LOSS_ITEM_ID = -1
   D.FROM_HOUSE_ID = LocationID
   If Not (Params Is Nothing) Then
      D.OrderBy = Params("ORDER_BY")
      D.OrderType = Params("ORDER_TYPE")
   Else
      D.OrderBy = 1
   End If
   D.OrderType = 1
   D.PARENT_FLAG = ParentFlag
   D.PIG_TYPE = PigType
   D.YEAR_SEQ_ID = YearID
   D.WEEK_NO = WeekNo
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLossItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set Pi = Partitems(Trim(Str(TempData.PIG_ID)))
         Call Cl.Add(Pi, Trim(Str(Pi.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Partitems = Nothing
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumIntakeByPigYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CIntake
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CIntake
Dim I As Long

   Set D = New CIntake
   Set Rs = New ADODB.Recordset
   
   D.INTAKE_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CIntake
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.UNIT_NAME)
'         C.ItemData(i) = TempData.UNIT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumIntakeByPigPartItemYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CIntake
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CIntake
Dim I As Long

   Set D = New CIntake
   Set Rs = New ADODB.Recordset
   
   D.INTAKE_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BATCH_ID = BatchID
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CIntake
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.UNIT_NAME)
'         C.ItemData(i) = TempData.UNIT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.PIG_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.YYYYMM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitInTakeType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("อาหาร")
   C.ItemData(1) = 1
   
   C.AddItem ("ยา+วัคซีน")
   C.ItemData(2) = 2
   
   C.AddItem ("อื่นๆ")
   C.ItemData(3) = 3
End Sub
Public Function InTakeTypeToString(ID As Long) As String
   If ID = 1 Then
      InTakeTypeToString = "อาหาร"
   ElseIf ID = 2 Then
      InTakeTypeToString = "ยา+วัคซีน"
   ElseIf ID = 3 Then
      InTakeTypeToString = "อื่นๆ"
   End If
End Function
Public Sub InitPoType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("PO สั่งซื้อวัตถุดิบ"))
   C.ItemData(1) = 1000

'   C.AddItem (MapText("PO สั่งซื้อวัสดุอุปกรณ์"))
'   C.ItemData(2) = 1001
'
'   C.AddItem (MapText("PO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์"))
'   C.ItemData(3) = 1002
'
'   C.AddItem (MapText("PO สั่งซื้อทั่วไป"))
'   C.ItemData(4) = 1003
End Sub
Public Sub LoadSumManagementExpense(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BatchID As Long = -1, Optional DepreciationFlag As String)
On Error GoTo ErrorHandler
Dim D As CParamItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CParamItem
Dim I As Long

   Set D = New CParamItem
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("PARAM_AREA", 15)
   Call D.SetFieldValue("BATCH_ID", BatchID)
   Call D.SetFieldValue("DEPRECIATION_FLAG", DepreciationFlag)
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CParamItem
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.GetFieldValue("YYYYMM"))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitPriceAdjustOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสสินค้า")
   C.ItemData(1) = 1
   
End Sub
Public Sub LoadSumUpdateAvg(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPriceAdjust
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPriceAdjust
Dim I As Long

   Set D = New CPriceAdjust
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPriceAdjust
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadExportPriceByHouseGroup(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PartGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long

   Set Cl = Nothing
   Set Cl = New Collection
   
   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.EXPORT_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PART_GROUP_ID = PartGroupID
   Call D.QueryData(59, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(59, Rs)
                     
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.DOCUMENT_DATE) & "-" & Trim(TempData.HOUSE_GROUP_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set D = Nothing
   Set Rs = Nothing
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumAgeWeight(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String, Optional DocTypeSet As String, Optional PigStatus As Long)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
Dim TempData2 As CDoItem
   
   Set Cl = Nothing
   Set Cl = New Collection
   
   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = "Y"
   D.COMMIT_FLAG = CommitFlag
   D.DocTypeSet = DocTypeSet
   D.PIG_STATUS = PigStatus
   Call D.QueryData(31, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(31, Rs)
      
      If Not (Cl Is Nothing) Then
         Set TempData2 = GetObject("CDoItem", Cl, Trim(TempData.PIG_AGE & "-" & Round(MyDiff(TempData.TOTAL_WEIGHT, TempData.ITEM_AMOUNT), 0)), False)
         If TempData2 Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.PIG_AGE & "-" & Round(MyDiff(TempData.TOTAL_WEIGHT, TempData.ITEM_AMOUNT), 0)))
         Else
            TempData2.ITEM_AMOUNT = TempData2.ITEM_AMOUNT + TempData.ITEM_AMOUNT
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set D = Nothing
   Set Rs = Nothing
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumAllCustom(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String, Optional DocTypeSet As String, Optional Cuscode As String, Optional PigStatus As Long)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
Dim TempData2 As CDoItem
   
   Set Cl = Nothing
   Set Cl = New Collection
   
   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.DocTypeSet = DocTypeSet
   D.CUSTOMER_CODE = Cuscode
   D.PIG_STATUS = PigStatus
   D.TAKE_KEY_ACCOUNT = "N"

   Call D.QueryData(50, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(50, Rs)
      
      If Not (Cl Is Nothing) Then
         Set TempData2 = GetObject("CDoItem", Cl, Trim(TempData.CUSTOMER_ID & "-" & TempData.CUSTOMER_CODE), False)
         If TempData2 Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.CUSTOMER_ID & "-" & TempData.CUSTOMER_CODE))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set D = Nothing
   Set Rs = Nothing
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctPigAge(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatus As Long)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
Dim TempData2 As CDoItem
   
   Set Cl = Nothing
   Set Cl = New Collection
   
   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = "Y"
   D.PIG_STATUS = PigStatus
   Call D.QueryData(48, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(48, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set D = Nothing
   Set Rs = Nothing
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumWeightAmountByPigAge(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PigStatus As Long)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
Dim TempData2 As CDoItem
   
   Set Cl = Nothing
   Set Cl = New Collection
   
   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = "Y"
   D.PIG_STATUS = PigStatus
   Call D.QueryData(49, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(49, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PIG_AGE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set D = Nothing
   Set Rs = Nothing
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadWeightAmountPriceByCus(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String, Optional DocTypeSet As String, Optional PigStatus As Long)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
Dim TempData2 As CDoItem
   
   Set Cl = Nothing
   Set Cl = New Collection
   
   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = "Y"
   D.COMMIT_FLAG = CommitFlag
   D.DocTypeSet = DocTypeSet
   D.PIG_STATUS = PigStatus
   Call D.QueryData(32, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(32, Rs)
      
      If Not (Cl Is Nothing) Then
         Set TempData2 = GetObject("CDoItem", Cl, Trim(TempData.CUSTOMER_ID & "-" & Round(MyDiff(TempData.TOTAL_WEIGHT, TempData.ITEM_AMOUNT), 0)), False)
         If TempData2 Is Nothing Then
            TempData.AVG_WEIGHT = Round(MyDiff(TempData.TOTAL_WEIGHT, TempData.ITEM_AMOUNT), 0)
            Call Cl.Add(TempData, Trim(TempData.CUSTOMER_ID & "-" & Round(MyDiff(TempData.TOTAL_WEIGHT, TempData.ITEM_AMOUNT), 0)))
         Else
            TempData2.ITEM_AMOUNT = TempData2.ITEM_AMOUNT + TempData.ITEM_AMOUNT
            TempData2.TOTAL_WEIGHT = TempData2.TOTAL_WEIGHT + TempData.TOTAL_WEIGHT
            TempData2.TOTAL_PRICE = TempData2.TOTAL_PRICE + TempData.TOTAL_PRICE
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set D = Nothing
   Set Rs = Nothing
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumWeightAmountPriceByCus(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String, Optional DocTypeSet As String, Optional PigStatus As Long)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
Dim TempData2 As CDoItem
   
   Set Cl = Nothing
   Set Cl = New Collection
   
   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PIG_FLAG = "Y"
   D.COMMIT_FLAG = CommitFlag
   D.DocTypeSet = DocTypeSet
   D.PIG_STATUS = PigStatus
   Call D.QueryData(33, Rs, itemcount)
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(33, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set D = Nothing
   Set Rs = Nothing
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInventoryDocSearch(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CInventoryDocSearch
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryDocSearch
Dim I As Long

   Set D = New CInventoryDocSearch
   Set Rs = New ADODB.Recordset
   
   D.INVENTORY_DOC_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CInventoryDocSearch
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.INVENTORY_DOC_ID)
         C.ItemData(I) = TempData.INVENTORY_DOC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.INVENTORY_DOC_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetPigStatusCustomerYYYYMM(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim Di As CDoItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
   
   Set Di = New CDoItem
   Set Rs = New ADODB.Recordset
   
    Di.DO_ITEM_ID = -1
   Di.FROM_DATE = FromDate
   Di.TO_DATE = ToDate
   Di.DOCUMENT_TYPE = DocumentType
   
   Call Di.QueryData(46, Rs, itemcount)
    
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(46, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PIG_STATUS & "-" & TempData.CUSTOMER_ID & "-" & TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
  
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub
Public Sub LoadPigImportAmount2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional PigType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PIG_TYPE = PigTypeToCode(PigType)
   Call D.QueryData(15, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPigExportAmount2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional PigType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
   
   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PIG_TYPE = PigTypeToCode(PigType)
   Call D.QueryData(23, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(23, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPigImportAmountDate2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional PigType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CImportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CImportItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CImportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PIG_TYPE = PigTypeToCode(PigType)
   Call D.QueryData(33, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CImportItem
      Call TempData.PopulateFromRS(33, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPigExportAmountDate2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional PigType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CExportItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExportItem
Dim I As Long
   
   Set D = New CExportItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PIG_FLAG = "Y"
   D.PIG_TYPE = PigTypeToCode(PigType)
   Call D.QueryData(64, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExportItem
      Call TempData.PopulateFromRS(64, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
