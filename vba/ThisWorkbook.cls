Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    Dim fileName
    IsLogined = False
    IsFinished = False
    
    HONAQ001Count = -1
    HONBQ001Count = -1
    HONAQ003Count = -1
    HONBQ003Count = -1
    
    If Sheets("모듈").OrderObject Is Nothing Then
        Set Sheets("모듈").OrderObject = New BizServer
        'Sheets("Sheet1").OrderObject.SetMainDir "C:\ExcelAPI"
    End If
    
    
'    fileName = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)) & "_체결내역.xlsx"
    
'    Sheets("체결내역 파일 보관").Range("B1").Value = fileName
    
End Sub


