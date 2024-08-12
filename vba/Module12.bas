Attribute VB_Name = "Module12"
Sub 매매계속()
Attribute 매매계속.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim iCount As Integer
'
' 매매계속 Macro


    WaitFor (3)
'시세중지
    Sheets("모듈").QuotePause_Click
    
    SendDataToMain GetHandleValue, MAIN_MACRO_NEXT, ""
        
'
    Sheets("가격수령").Select
    
 '   ActiveWorkbook.Save
    
    Range("C49:OL50").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C59").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Range("C15:OL16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
 '   ActiveWorkbook.Save
    
    
  '  Sheets("시뮬레이션").Select
   ' Range("d4").Select
   ' Application.CutCopyMode = False
   ' Selection.Copy
  '  Range("f4").Select
   ' Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
     '   :=False, Transpose:=False
    
      
   ' Sheets("주문체결보관").Select
    'Range("A11").Select
   ' Selection.Copy
   ' Range("A12").Select
   ' Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
   '     :=False, Transpose:=False
        

   
    
    Sheets("시뮬레이션").Select
    
'    ActiveWorkbook.Save
    Range("a5:GY93").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=7
    Range("a95").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
'    Range("F101").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("조건입력").Select
'    Range("E7").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
   
 '   ActiveWorkbook.Save
    
'    Sheets("시뮬레이션").Select
    
'    Range("F102").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("조건입력").Select
'    Range("N7").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
 '       :=False, Transpose:=False
    
'    Range("G115").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("조건입력").Select
'    Range("E6").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False

'    Sheets("시뮬레이션").Select
'    ActiveWorkbook.Save
'    Range("F113").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("조건입력").Select
'    Range("W13").Select
 '   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
 '       :=False, Transpose:=False
        
 '   ActiveWorkbook.Save
    Sheets("시뮬레이션").Select
    Range("G120").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Range("F30").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("C140").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("C146").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C54").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("C150").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C58").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("C157").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Range("C65").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("C160").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C68").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("H104:GY104").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Range("H12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H163:GY163").Select
    Selection.Copy
    Range("H71").Select
    ActiveSheet.Paste
    
'    Range("H107:GY107").Select
'    Application.CutCopyMode = False
'    Selection.Copy
    
'    Range("H16").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
      
    Range("H109:GY109").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H18").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H116:GY116").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'    Range("H120:GY120").Select
'    Application.CutCopyMode = False
'    Selection.Copy
    
'    Range("H27").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    
    Range("H124:GY124").Select
    Selection.Copy
    Range("H33").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H140:GY140").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H146:GY146").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H54").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H150:GY150").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H58").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
  
   
    Range("H160:GY160").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H68").Select
    ActiveSheet.Paste
    
    Range("H163:GY163").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H71").Select
    ActiveSheet.Paste
    
 
'    Range("N168:N176").Select
'    Application.CutCopyMode = False
'    Selection.Copy
    
'    Range("M78").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
 
   ' Range("f4").Select
   ' Application.CutCopyMode = False
   ' Selection.Copy
   ' Range("D3").Select
   ' Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
   '     :=False, Transpose:=False
   

    Range("H69").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("H69").Select
    Selection.Copy
    Range("I69:GY69").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'    ActiveWorkbook.Save
     
     
             '1. "주문체결보관"시트에서 가격수령란에 가격등 자료를 다운 받은후 다음을 수행
     
    
    Sheets("가격수령").Select
    
    Rows("6:10").Select
    Selection.Copy
    Range("A86").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'    sleep (3000)
'    ActiveWorkbook.Save
    Range("C40:GT44").Select
    Selection.Copy
    Sheets("주문체결보관").Select
    Range("C6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
 
'    Sheets("가격수령").Select
'     ActiveWorkbook.Save
'    Range("GV6:HC12").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("주문체결보관").Select
'    Range("GV6").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'     ActiveWorkbook.Save
'    Sheets("가격수령").Select
'     ActiveWorkbook.Save
'    Range("HE6:AMZ12").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("주문체결보관").Select
'    Range("HE6").Select
 '   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
 '       :=False, Transpose:=False
     
'    ActiveWorkbook.Save
    Range("C5:GT5").Select
    Selection.Copy
    Sheets("시뮬레이션").Select
    Range("H15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
'    ActiveWorkbook.Save
    
    Sheets("주문체결보관").Select
     ActiveWorkbook.Save
    Range("C13:GT13").Select
    Selection.Copy
    Sheets("시뮬레이션").Select
    Range("H15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
'    ActiveWorkbook.Save
    
    Sheets("주문체결보관").Select
     ActiveWorkbook.Save
    Range("C11:GT12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C27").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C18:GT18").Select
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
 '   sleep (3000)
 '   ActiveWorkbook.Save
    
    Dim nCheck As Variant
    nCheck = Range("A5").Value
    If nCheck > 0 Then
        매크로1
    Else
        매크로2
    End If
    

    Sheets("시뮬레이션").Select
    Range("h190").Select
    Range("a6").Select
    
    
    If OrderGubun = "HINT" Then
        StockHONAQ003Search
    ElseIf OrderGubun = "SK" Then
        Sheets("모듈").btnSKStock96013_Click
    ElseIf OrderGubun = "IBKS" Then
        Sheets("모듈").btnIBKSTR1211_Click
    End If
    
    ActiveWorkbook.Save

    '시세시작
    If IsCancelPress = False Then
        Sheets("모듈").QuoteStart_Click
    End If

    HONAQ001Count = -1
    HONBQ001Count = -1
    HONAQ003Count = -1
    HONBQ003Count = -1

End Sub
   
