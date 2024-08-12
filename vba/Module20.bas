Attribute VB_Name = "Module20"
Sub 첫매매()
Attribute 첫매매.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim iCount As Integer
'
' 첫매매 Macro
'

'시세중지
''잔고엑셀초기화
    
    Sheets("모듈").QuotePause_Click
    WaitFor (1)
    
    SendDataToMain GetHandleValue, MAIN_MACRO_START, ""
'
   '            1. 가격수령 시트의 가격란에 가격을 입력한 후에 다음을 수행
    Sheets("가격수령").Select
    
    Rows("6:10").Select
    Selection.Copy
    Range("A86").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("C39:GT39").Select
    Selection.Copy
    Sheets("주문체결보관").Select
    Range("C31:GT31").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
   
   
   
    Sheets("가격수령").Select
   
    ActiveWorkbook.Save
     
    Range("C40:GT44").Select
    Selection.Copy
    Sheets("주문체결보관").Select
    Range("C6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'    Sheets("가격수령").Select
'    Range("GV6:HC12").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("주문체결보관").Select
'    Range("GV6").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Sheets("가격수령").Select
'    Range("HE6:AMZ12").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("주문체결보관").Select
'    Range("HE6").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'       :=False, Transpose:=False
    
    ActiveWorkbook.Save

   
    Sheets("가격수령").Select
    Range("C25").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
    Range("C26").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C25:C26").Select
    Selection.Copy
'    Range("D25:GT26,GV25:HC26,HE25:AMZ26").Select
    Range("D25:OL26").Select
    Range("d25").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Range("C59").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
    Range("C60").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C59:C60").Select
    Selection.Copy
'    Range("D25:GT26,GV25:HC26,HE25:AMZ26").Select
    Range("D59:gt60").Select
    Range("d59").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Sheets("주문체결보관").Select
    Rows("63:50000").Select
    Selection.Delete Shift:=xlUp
    Range("A47").Select
    
    Sheets("조건입력").Select
'    Range("E6").Select
'    ActiveCell.FormulaR1C1 = "=0"
'    Range("E7").Select
'    ActiveCell.FormulaR1C1 = "=0"
'    Range("N7").Select
'    ActiveCell.FormulaR1C1 = "=0"
    Range("W13").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("W14:X14").Select
    
    Sheets("종목분석").Select
    
    ActiveWorkbook.Save
   
    Sheets("주문체결보관").Select
    Range("C6:GT6").Select
    Selection.Copy
    Sheets("시뮬레이션").Select
    Range("H15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("d97").Select
    ActiveCell.FormulaR1C1 = "=0"
    
    
    
    Range("H192").Select
    ActiveCell.FormulaR1C1 = "=R[-5]C*R[21]C/SUMPRODUCT(R3C8:R3C207,R29C8:R29C207)"
    Range("H192").Select
    Selection.Copy
    Range("I192:GY192").Select
    ActiveSheet.Paste
    
    
    Range("H25").Select
    ActiveCell.FormulaR1C1 = _
        "=ROUND(R1C2*100000000*R[-17]C,0)-INT(R[-19]C*R[-18]C*(1+R2C7))"
    Range("H25").Select
    Selection.Copy
    Range("I25:GY25").Select
    ActiveSheet.Paste
 
 
    Range("C65").Select
    ActiveCell.FormulaR1C1 = "=0"
    
    Range("h68").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("h68").Select
    Selection.Copy
    Range("i68:gy68").Select
    ActiveSheet.Paste

    Range("F7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("F10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    

    Range("H12").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-6]C"
    Range("H12").Select
    Selection.Copy
    Range("I12:GY12").Select
    ActiveSheet.Paste
    
    Range("H15:GY15").Select
    Application.CutCopyMode = False
    Selection.Copy
    
'    Range("H7").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    Range("H18").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H33").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Range("f9").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Range("c33").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'    Range("H16").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "=0"
'    Range("H16").Select
'    Selection.Copy
'    Range("I16:GY16").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
        
'    Range("c27").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "=0"
'    Range("c27").Select
'    Selection.Copy
'    Range("h27:GY27").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    
    Range("F30").Select
    ActiveCell.FormulaR1C1 = "=0" _
'        "=R[-29]C[-4]*100000000-SUMPRODUCT(R[24]C[2]:R[24]C[201],R[-15]C[2]:R[-15]C[201])"
        
    Range("H35:GY35").Select
    Selection.Copy
    Range("H36").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
  
    Range("c35").Select
    Selection.Copy
    Range("c36").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("C48").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    
    Range("c54").Select
    ActiveCell.FormulaR1C1 = "=0"
    
 
    Range("c68").Select
    ActiveCell.FormulaR1C1 = "=0"
 
    
    Range("C58").Select
    ActiveCell.FormulaR1C1 = "=0"
    
    
    Range("C65").Select
    ActiveCell.FormulaR1C1 = "=0"
  
    
'    Range("M78").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "=0"
    
'    Range("M79").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "=0"
    
'    Range("M81").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "=0"
    
'    Range("M82").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "=0"
             
'     Range("M84").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "=0"
    
'    Range("M85").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "=0"
   
    
    Range("H28:GY28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H29").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
     
    Range("H58").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("H58").Select
    Selection.Copy
    Range("I58:GY58").Select
    ActiveSheet.Paste
    
    Range("H48").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("H48").Select
    Selection.Copy
    Range("I48:GY48").Select
    ActiveSheet.Paste
    
    Range("H54").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
'   ActiveCell.FormulaR1C1 = "=ROUND(R[+3]C,0)"
    Range("H54").Select
    Selection.Copy
    Range("I54:GY54").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
     
    
    Range("H71").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("H71").Select
    Selection.Copy
    Range("I71:GY71").Select
    ActiveSheet.Paste
     
    Range("G117").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("G118").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("G119").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("f123").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("f124").Select
    ActiveCell.FormulaR1C1 = "=0"
   
    Range("f125").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("f126").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("g124").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("g125").Select
    ActiveCell.FormulaR1C1 = "=0"
    ActiveCell.FormulaR1C1 = "=0"
    Range("f130").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("g132").Select
    ActiveCell.FormulaR1C1 = "=0"
   
 
    Range("g137").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("c146").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("c147").Select
    ActiveCell.FormulaR1C1 = "=0"
    
   
    Range("d149").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("c151").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("c160").Select
    ActiveCell.FormulaR1C1 = "=0"
'    Range("k168").Select
'    ActiveCell.FormulaR1C1 = "=0"
   
    Range("a5:GY93").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=7
    Range("a95").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
      Range("h106").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("h106").Select
    Selection.Copy
    Range("i106:gy106").Select
    ActiveSheet.Paste


    Range("h120").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("h120").Select
    Selection.Copy
    Range("i120:gy120").Select
    ActiveSheet.Paste

    Range("h147").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("h147").Select
    Selection.Copy
    Range("i147:gy147").Select
    ActiveSheet.Paste
    
    Range("h128").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("h128").Select
    Selection.Copy
    Range("i128:gy128").Select
    ActiveSheet.Paste


 '   Range("h117").Select
 '   ActiveCell.FormulaR1C1 = "=0"
 '   Range("h117").Select
 '   Selection.Copy
 '   Range("i117:gy117").Select
 '   ActiveSheet.Paste

    Range("h140").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("h140").Select
    Selection.Copy
    Range("i140:gy140").Select
    ActiveSheet.Paste


    Range("h141").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("h141").Select
    Selection.Copy
    Range("i141:gy141").Select
    ActiveSheet.Paste

    Range("h154").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("h154").Select
    Selection.Copy
    Range("i154:gy154").Select
    ActiveSheet.Paste
    
    Range("H134").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("H135").Select
    ActiveCell.FormulaR1C1 = "=0"
    
    Range("H134:h135").Select
    Selection.Copy
    Range("I134:GY134").Select
    ActiveSheet.Paste
     
    Range("HA5:Hw5").Select
    Selection.Copy
    ActiveWindow.SmallScroll ToRight:=-6
    Range("HA6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("HA7:Hw10000").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("HA6:Hw6").Select
    Selection.Copy
    ActiveWindow.SmallScroll ToRight:=-5
    Range("HA7:Hw7").Select
    Selection.Insert Shift:=xlDown
    ActiveWindow.SmallScroll ToRight:=-6
        
    Range("HA6:Hw321").Select
    ActiveWindow.SmallScroll ToRight:=12
    Selection.ClearContents
      
    Range("d143").Select
    ActiveCell.FormulaR1C1 = "=0"
     
    Range("H25:GY25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

   
    Range("H156").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=0"
    Range("H157").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("H156:H157").Select
    Selection.AutoFill Destination:=Range("H156:GY157"), Type:=xlFillDefault
    Range("H156:GY157").Select
                
                
    Range("H165").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("H165").Select
    Selection.AutoFill Destination:=Range("H165:GY165"), Type:=xlFillDefault
    Range("H165:GY165").Select
                
    Range("H128").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("H128").Select
    Selection.AutoFill Destination:=Range("H128:GY128"), Type:=xlFillDefault
    Range("H128:GY128").Select
    
    Range("H187:GY187").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H178").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'sleep (3000)
    WaitFor (3)
    ActiveWorkbook.Save
            
                
    Sheets("주문체결보관").Select
    
    ActiveWorkbook.Save
  
    Range("C13:GT13").Select
    Selection.Copy
    Sheets("시뮬레이션").Select
    Range("H15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ActiveWorkbook.Save

   
    Sheets("주문체결보관").Select
    Range("C11:GT12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C27").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C18:gt18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C19:GT19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C17").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
   
    ActiveWorkbook.Save

    Range("c14:gt17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("가격수령").Select
    Range("c63").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Range("C72:OL75").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("c29").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    ActiveWorkbook.Save
 
     
  '2.여기까지 수행후 주문란의 주문을 실행하고 체결 결과를  기다린 후
  
  '3.미체결은 취소하고 체결후의 잔고주식수와 평균단가를를 가격수령시트에 입력하고 다음을 수행

    Dim iCancelSec As Integer
    iCancelSec = Int(Sheets("모듈").Cells(3, 5).Value)

'주문
    '주식
    If OrderGubun = "HINT" Then
        WaitFor (1)
        Sheets("모듈").StockBatchBuy_Click
        WaitFor (1)
        Sheets("모듈").StockBatchSell_Click
    ElseIf OrderGubun = "SK" Then
        WaitFor (1)
        Sheets("모듈").btnSKStockBuy_Click
        WaitFor (1)
        Sheets("모듈").btnSKStockSell_Click
    ElseIf OrderGubun = "IBKS" Then
        WaitFor (1)
        Sheets("모듈").btnIBKSStockBuy_Click
        WaitFor (1)
        Sheets("모듈").btnIBKSStockSell_Click
    End If
    


    WaitFor (iCancelSec)
'미체결내역조회
'취소(주식)
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

    Sheets("가격수령").Select
    ActiveWorkbook.Save
 
     
    Range("C53:GT56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("주문체결보관").Select
    Range("C20").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
   
    ActiveWorkbook.Save

    Range("C26:GT26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("시뮬레이션").Select
    Range("H69").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
     :=False, Transpose:=False
    
    ActiveWorkbook.Save

    Sheets("주문체결보관").Select
    ActiveWorkbook.Save
 
    Range("C29:GT29").Select
    Selection.Copy
    Sheets("시뮬레이션").Select
    Range("H15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      
    Sheets("주문체결보관").Select
    
    Rows("6:10").Select
    Selection.Copy
    Range("A31").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
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

    Sheets("시뮬레이션").Select
    Range("d97").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("d125").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("f137").Select
    ActiveCell.FormulaR1C1 = "=0"
    Range("f138").Select
    ActiveCell.FormulaR1C1 = "=0"
   
   
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
    
    Range("H25:GY25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'    Range("H8:GY8").Select
'    Application.CutCopyMode = False
'    Selection.Copy
 '   Range("H8").Select
 '   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
 '       :=False, Transpose:=False
  
   
    Range("H6:GY6").Select
    Selection.Copy
    Range("H37").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
  
    
    Range("a5:GY93").Select
    Selection.Copy
    Range("A95").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H15:GY15").Select
    Selection.Copy
    Range("H7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H25:GY25").Select
    Selection.Copy
    Range("H24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Range("a6").Select
    
    Range("H3").Select

   
    ActiveWorkbook.Save

'시세시작
    Sheets("모듈").QuoteStart_Click

    HONAQ001Count = -1
    HONBQ001Count = -1
    HONAQ003Count = -1
    HONBQ003Count = -1
    
End Sub
    
