Attribute VB_Name = "Sheet7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "btnStockHONAQ003_, 1, 0, MSForms, CommandButton"
Private Sub btnStockHONAQ003__Click()
    '주식:주문체결/전체 조회(96012)

    Dim nPoint As Long
    
    Sheets("체결내역").Range("A3:L2000").ClearContents
    isSelectHONAQ003 = True
    
    nPoint = Sheets("모듈").WapiCommunityMngCtrl.InitTR("96012")
    
    If nPoint > 0 Then
        Call Sheets("모듈").WapiCommunityMngCtrl.SetTrData(nPoint, "96012.계좌번호", Sheets("모듈").Cells(4, 9).Value, 0)
        Call Sheets("모듈").WapiCommunityMngCtrl.SetTrData(nPoint, "96012.비밀번호", Sheets("모듈").Cells(5, 9).Value, 0)
        Call Sheets("모듈").WapiCommunityMngCtrl.SetTrData(nPoint, "96012.종목코드", "", 0)
        Call Sheets("모듈").WapiCommunityMngCtrl.SetTrData(nPoint, "96012.매매구분", "0", 0)   '0:전체, 1:매도, 2:매수
        Call Sheets("모듈").WapiCommunityMngCtrl.SetTrData(nPoint, "96012.주문대리인구분", "0", 0)   '0:일반, 1:주문대리인
        Call Sheets("모듈").WapiCommunityMngCtrl.SetTrData(nPoint, "96012.대표계좌번호", "", 0) '주문대리인구분 1로 선택했을 경우 주문대리인 계좌번호 입력
        Call Sheets("모듈").WapiCommunityMngCtrl.SendTr("96012", 0)
    End If
    
End Sub
