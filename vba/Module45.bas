Attribute VB_Name = "Module45"
Public Sub 매크로1()
Attribute 매크로1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로1 매크로
'

'
      
             
  
    Sheets("주문체결보관").Select
    
'    ActiveWorkbook.Save
    
    Range("c14:gt17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("가격수령").Select
    Range("c63").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
'    ActiveWorkbook.Save
             
             
    Range("c72:OL75").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("c29").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
 '   ActiveWorkbook.Save
                        
             
             '2.여기까지 수행후 가격수령시트의 주문을 실행하고 체결 결과를  기다린 후
             
             '3.미체결은 취소하고 체결결과를 가격수령시트의 결과란에 입력하고 다음을 수행
    Dim iCancelSec As Integer
    iCancelSec = Int(Sheets("모듈").Cells(3, 5).Value)
    
    '주문(주식)
    If OrderGubun = "HINT" Then
        Sheets("모듈").StockBatchBuy_Click
        WaitFor (1)
        Sheets("모듈").StockBatchSell_Click
    ElseIf OrderGubun = "SK" Then
        Sheets("모듈").btnSKStockBuy_Click
        WaitFor (1)
        Sheets("모듈").btnSKStockSell_Click
    ElseIf OrderGubun = "IBKS" Then
        Sheets("모듈").btnIBKSStockBuy_Click
        WaitFor (1)
        Sheets("모듈").btnIBKSStockSell_Click
    End If
    
    WaitFor (iCancelSec)
    
    '미체결내역조회
    '취소
    Do While True
    '   주식취소
        If OrderGubun = "HINT" Then
            Sheets("모듈").btnStockHONAQ003_Click
        ElseIf OrderGubun = "SK" Then
            Sheets("모듈").btnSKStock96013_Click
        ElseIf OrderGubun = "IBKS" Then
            Sheets("모듈").btnIBKSTR1211_1_Click
        End If
        
        iCount = 0
        Do While iCount < 30
            WaitFor (1)
            If (HONAQ003Count >= 0) Then
                Exit Do
            End If
            iCount = iCount + 1
        Loop
        
        WaitFor (1)
        
    '   취소전송
        If (HONAQ003Count > 0) Then
            If OrderGubun = "HINT" Then
                Sheets("모듈").StockBatchCancel_Click
            ElseIf OrderGubun = "SK" Then
                Sheets("모듈").btnSKStockCancel_Click
            ElseIf OrderGubun = "IBKS" Then
                Sheets("모듈").btnIBKSStockCancel_Click
            End If
            
        End If
        
        WaitFor (1)
        ' 선물을 운용 할 때 HONBQ003Count도 따로 체크 해야함.
        If (HONAQ003Count <= 0) Then
            Exit Do
        End If
           
    Loop

    '잔고조회(주식)
    If OrderGubun = "HINT" Then
        Sheets("모듈").btnStockHONAQ001_Click
    ElseIf OrderGubun = "SK" Then
        Sheets("모듈").btnSKStock96007_Click
    ElseIf OrderGubun = "IBKS" Then
        Sheets("모듈").btnIBKSTR1221_Click
    End If
    
    iCount = 0
    Do While iCount < 30
        WaitFor (1)
        If HONAQ001Count >= 0 Then
            Exit Do
        End If
        iCount = iCount + 1
    Loop
             
    '잔고조회(선물)
'    Sheets("모듈").btnFOHONBQ001_Click
'    iCount = 0
'    Do While iCount < 30
'        WaitFor (1)
'        If HONBQ001Count >= 0 Then
'            Exit Do
'        End If
'        iCount = iCount + 1
'    Loop
             

   
                 
    Sheets("가격수령").Select
   
 '   ActiveWorkbook.Save
    Range("c53:gt56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("주문체결보관").Select
    Range("c20").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
 '   ActiveWorkbook.Save
    
    Sheets("주문체결보관").Select
    Range("C26:GT26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("시뮬레이션").Select
    Range("H69").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
 '   ActiveWorkbook.Save
    
    Sheets("주문체결보관").Select
    ActiveWorkbook.Save
    Rows("6:10").Select
    Selection.Copy
    Range("A31").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
 '   ActiveWorkbook.Save
    
    Rows("41:52").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=9
    Rows("54:54").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Rows("67:67").Select
    Selection.Insert Shift:=xlDown
     
    Sheets("주문체결보관").Select
    ActiveWorkbook.Save
    Range("C29:GT29").Select
    Selection.Copy
    Sheets("시뮬레이션").Select
    Range("H15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
'    ActiveWorkbook.Save

    Sheets("시뮬레이션").Select
    Range("HA5:Hw5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("HA6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("HA7:Hw7").Select
    Selection.Insert Shift:=xlDown
          
'    Range("M77,M80,M83,M86").Select
'    Range("M86").Activate
'    Selection.NumberFormatLocal = "#,##0_ "

   
 '   ActiveWorkbook.Save

End Sub
