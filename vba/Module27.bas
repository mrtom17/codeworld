Attribute VB_Name = "Module27"
Sub 연속매매()
Attribute 연속매매.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 연속매매 Macro
'
    Dim sEndTime As String
    Dim iNextSec As Integer
    
    SendDataToMain GetHandleValue, MAIN_MACRO_CONTINUE, ""
    
    sEndTime = Sheets("모듈").Cells(1, 5).Value
    iNextSec = Int(Sheets("모듈").Cells(2, 5).Value)

    IsCancelPress = False
    Application.OnTime TimeValue(sEndTime), "Finished"
    For y = 1 To 1800
       
        If IsCancelPress Then
            Exit For
        End If
        
        Call 매매계속
        
        WaitFor (iNextSec)
    Next y

    'MsgBox "finish"

End Sub
