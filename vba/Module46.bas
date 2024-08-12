Attribute VB_Name = "Module46"
Sub 매크로2()
Attribute 매크로2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로2 매크로
'

    Sheets("주문체결보관").Select
    
 '   ActiveWorkbook.Save
    Range("c14:gt17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("가격수령").Select
    Range("c63").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("c72:ol75").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("c29").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
 '   ActiveWorkbook.Save
    
    Range("C53:GT56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("주문체결보관").Select
    Range("C20").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
 '   ActiveWorkbook.Save
   
   Sheets("주문체결보관").Select
   Range("H3").Select
   
End Sub
