Attribute VB_Name = "Module3"
Sub 매크로3()
Attribute 매크로3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로3 매크로
'

'
    Range("C39:GT39").Select
    Selection.Copy
    Sheets("주문체결보관").Select
    Range("C31:GT31").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
