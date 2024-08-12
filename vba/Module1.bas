Attribute VB_Name = "Module1"
Option Explicit

Public IsLogined As Boolean
Public IsFinished As Boolean
Public IsCancelPress As Boolean

Public AccountSeq As Integer
Public AccountSeqFO As Integer

Public AccountList() As String
Public AccountListFO() As String

Public UserID As String
Public UserPass As String
Public AccountPass As String
Public ReplayGubun As String
Public OrderGubun As String

Public CurrentSheetGubun As String

Public QuoteStockHStartPos As Integer
Public QuoteStockVStartPos As Integer
Public QuoteStockItemCount As Integer

Public QuoteFutureHStartPos As Integer
Public QuoteFutureVStartPos As Integer
Public QuoteFutureItemCount As Integer


Public HONAQ001Data() As String
Public HONAQ003Data() As String
Public HONBQ001Data() As String
Public HONBQ003Data() As String

Public isSelectHONAQ003 As Boolean

Public HONAQ001Count As Integer
Public HONAQ003Count As Integer
Public HONBQ001Count As Integer
Public HONBQ003Count As Integer

Public HONAQ001Count_T As Integer
Public HONAQ003Count_T As Integer


Public OldAccountStock As String
Public OldAccountStock1 As String
Public OldAccountFO As String
Public OldAccountFO1 As String


Public screenUpdateState As Boolean
Public statusBarState As Boolean
Public calcState As Boolean
Public eventsState As Boolean
Public displayPageBreakState As Boolean

#If VBA7 Then
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lpData As COPYDATASTRUCT) As LongPtr
#Else
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lpData As COPYDATASTRUCT) As Long
#End If

Public Type COPYDATASTRUCT
    dwData As LongPtr ' 변경된 부분
    cbData As Long
    lpData As LongPtr ' 변경된 부분
End Type

Public Const WM_COPYDATA As Long = &H4A
Public Const WM_USER As Long = &H400

' 계좌정보
Public Const MAIN_ACCOUNT_LIST As Long = WM_USER + 1200
' 주식 잔고평가내역
Public Const MAIN_BUSI_RECV_HONAQ001_H As Long = WM_USER + 1201
Public Const MAIN_BUSI_RECV_HONAQ001 As Long = WM_USER + 1202
' 주식 채결내역
Public Const MAIN_BUSI_RECV_HONAQ003 As Long = WM_USER + 1203

' Notice(실시간 매매 노티 수신)
Public Const MAIN_NOTICE As Long = WM_USER + 1204

Public Const MAIN_SECURITY_READY As Long = WM_USER + 1206               '증권사 준비완료
Public Const MAIN_MACRO_START As Long = WM_USER + 1207                  '매크로 첫매매 처리완료
Public Const MAIN_MACRO_NEXT As Long = WM_USER + 1208                   '매크로 매매계속 처리 완료
Public Const MAIN_MACRO_CONTINUE As Long = WM_USER + 1209               '매크로 연속매매 처리 완료

Public Sub SendDataToMain(hwnd As LongPtr, ByVal dwData As Long, ByVal message As String)
    Dim cds As COPYDATASTRUCT
    Dim result As LongPtr ' 64비트 호환을 위해 수정
    
    ' COPYDATASTRUCT 설정
    With cds
        .dwData = dwData ' 사용자 정의 데이터를 위한 값
        .cbData = Len(message) * 2 ' 데이터 길이 (+1은 널 종료 문자를 위해)
        .lpData = StrPtr(message) ' 메시지의 포인터
    End With
    
    ' 데이터 보내기
    result = SendMessage(hwnd, WM_COPYDATA, hwnd, cds)
'    If result = 0 Then
'        MsgBox "Data transfer failed.", vbCritical
'    Else
'        MsgBox "Data sent successfully!", vbInformation
'    End If
End Sub

Public Function GetHandleValue() As LongPtr
    Dim handleVariant As Variant
    Dim handleValue As LongPtr
    
    ' OrderObject의 GetMainHandle 메서드를 호출하여 핸들 값을 얻음
    handleVariant = Sheets("모듈").OrderObject.GetMainHandle()
    
    ' Variant 타입에서 LongPtr 타입으로 변환
    handleValue = handleVariant
    
    ' 핸들 값을 반환
    GetHandleValue = handleValue
End Function

Public Sub Finished()
    IsCancelPress = True
    IsFinished = True
    'Sheets("모듈").QuoteStop_Click
End Sub

Public Sub CreateBizObject()
    
    If Sheets("모듈").OrderObject Is Nothing Then
        Set Sheets("모듈").OrderObject = New BizServer
        Sheets("모듈").OrderObject.SetMainDir "C:\ExcelAPI"

        AccountSeq = 0
    End If

    Sheets("모듈").OrderObject.SetMainDir "C:\ExcelAPI"

    AccountSeq = 0

'    If Sheets("모듈").WapiCommunityMngCtrl Is Nothing Then
'        Set Sheets("모듈").WapiCommunityMngCtrl = New WapiCommunity
'    End If
'
'    If Sheets("모듈").WapiConnectMngCtrl Is Nothing Then
'        Set Sheets("모듈").WapiConnectMngCtrl = New WapiConnectMng
'    End If

End Sub

Public Sub LoginProcess()
    CreateBizObject
    
'    Application.ScreenUpdating = False
'    'Application.DisplayStatusBar = False
'    Application.Calculation = xlCalculationManual
'    'Application.EnableEvents = False
'    'ActiveSheet.DisplayPageBreaks = False
    
'    Application.ScreenUpdating = screenUpdateState
'    'Application.DisplayStatusBar = statusBarState
'    'Application.Calculation = calcState
'    'Application.EnableEvents = eventsState
'    'ActiveSheet.DisplayPageBreaks = displayPageBreakState
    
    Sheets("모듈").OrderObject.LoginController UserID, UserPass, ReplayGubun
    

End Sub

Public Sub WaitFor(NumOfSeconds As Single)

    Dim SngSec As Single
    SngSec = Timer + NumOfSeconds

    Do While Timer < SngSec
        DoEvents
    Loop

End Sub

Public Sub StockHONAQ003Search()
    Dim sIssueCode As String     '계좌번호
    Dim sPassword As String     '비밀번호
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value
    
    sIssueCode = AccountList(0, 0)
    sPassword = AccountPass
    
    CreateBizObject
    
    HONAQ003Count = -1
    isSelectHONAQ003 = False
    
    Sheets("모듈").OrderObject.SendReceiveEx sIssueCode, sPassword, "HONAQ003", "0"   '체결내역조회 조회구분 0:전체(체결,미체결), 1:체결, 2:미체결
End Sub

Public Sub ParseJSONToSet(JsonString As String)
    Dim ParsedJson As Object ' late binding을 사용하여 Scripting.Dictionary의 인스턴스 생성
    
    Set ParsedJson = JsonConverter.ParseJson(JsonString)
    
    If ParsedJson.Exists("ExType") Then Sheets("모듈").Cells(4, 2).Value = ParsedJson("ExType")
    If ParsedJson.Exists("LoginID") Then Sheets("모듈").Cells(1, 2).Value = ParsedJson("LoginID")
    If ParsedJson.Exists("LoginPass") Then Sheets("모듈").Cells(2, 2).Value = ParsedJson("LoginPass")
    If ParsedJson.Exists("AccountPass") Then Sheets("모듈").Cells(3, 2).Value = ParsedJson("AccountPass")
    If ParsedJson.Exists("OrderType") Then Sheets("모듈").Cells(4, 5).Value = ParsedJson("OrderType")
    If ParsedJson.Exists("FinishTime") Then Sheets("모듈").Cells(1, 5).Value = ParsedJson("FinishTime")
    If ParsedJson.Exists("ContinueWait") Then Sheets("모듈").Cells(2, 5).Value = ParsedJson("ContinueWait")
    If ParsedJson.Exists("OrderWait") Then Sheets("모듈").Cells(3, 5).Value = ParsedJson("OrderWait")
        
End Sub

Public Sub CallQuote(pnMarketType As Integer, psItemSheetName As String, psModuleSheetName As String, pnSheetNum As Integer, _
                     pnItemY As Integer, pnItemX As Integer, _
                     pnCellY As Integer, pnCellX As Integer, pnTotalItemCount As Integer, pnCallMaxCount As Integer, _
                     Optional pnDirection As Integer = 0, Optional pnTimer As Integer = 0, Optional pnTimerType As Integer = 0, _
                     Optional pnWaitTime As Single = 0.1)
'Parameters
'    pnMarketType As String       '시장 구분(1:장내종목(코스피, 코스닥), 2:선물옵션)
'    psItemSheetName As String    '종목 리스트가 있는 엑셀 시트 명
'    psModuleSheetName As String  '모듈이 있는 엑셀 시트 명
'    pnSheetNum As Integer        '종목을 받는 엑셀 시트의 인덱스 번호
'    pnItemY As Integer           '세로 종목 위치
'    pnItemX As Integer           '가로 종목 위치
'    pnCellY As Integer           '세로 시세 시작 위치
'    pnCellX As Integer           '가로 시세 시작 위치
'    pnTotalItemCount As Integer  '총 종목수
'    pnCallMaxCount As Integer    '분할 호출 개수 : 최대 50종목까지만 호출 할 수 있음
'    pnDirection As Integer       '방향 : 0(X축), 1(Y축)-종목이 X축 방향, Y축 방향으로 위치해 있느냐를 선택하는 옵션
'    pnTimer As Integer           '타이머 : n(n초에 한번씩 시세를 뿌려줌)
'    pnTimerType As Integer       '타이머구분 : 0(시세를 호출한 시점 부터 시작), 1(정시를 기준으로 처리-시작시간이 3초에 시작된경우 5초에 처리됨)
'    pnWaitTime As Single         '시세 호출 후 대기 시간

    Dim nCol As Integer             '루프 변수
    Dim nCallItemCount As Integer   '호출한 종목 수
    Dim sItemCodes As String        '종목 리스트 ","로 구분한다.
    Dim sOrgCode As String          '주식 일때만 사용됨


    '여러번 호출 될 수 있기 때문에 여기에 있으면 안됨.
    'Application.Calculation = xlCalculationManual


    '종목 리스트 초기화
    sItemCodes = ""
    nCallItemCount = 0
    For nCol = 1 To pnTotalItemCount
        '종목 리스트 만들기
        '종목이 비어 있는지 확인
        If Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value <> "" Then
        '종목코드 값이 있는 경우
            If pnMarketType = 1 Then
            '주식
                sOrgCode = Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value
                sItemCodes = sItemCodes & Right(sOrgCode, Len(sOrgCode) - 1) & ","
            ElseIf pnMarketType = 2 Then
            '선물옵션
                sItemCodes = sItemCodes & Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value & ","
            End If
            
            nCallItemCount = nCallItemCount + 1
        Else
        '종목코드가 비어 있을 때
            '종목코드가 비어 있고 sItemCodes에 값이 존재 하는 경우
            If nCallItemCount > 0 Then
                '종목 리스트 마지막에 ","제거
                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
                
                '시세 호출
                Sheets(psModuleSheetName).OrderObject.SetQuote sItemCodes, pnSheetNum, pnCellY, nCol - nCallItemCount - 1 + pnCellX, _
                                                               pnDirection, pnTimer, pnTimerType
                
                '시세 호출 후 종목 리스트 초기화
                sItemCodes = ""
                nCallItemCount = 0
    
                'Delay 없는  연속 처리 금지
                WaitFor (pnWaitTime)
                
            End If
            GoTo Continue
        
        End If

        
        '호출 할 종목 체크
        If (nCallItemCount > 0) And (((nCol Mod pnCallMaxCount) = 0) Or (nCol = pnTotalItemCount)) Then
            '종목 리스트 마지막에 ","제거
            sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
            
            '시세 호출
            If (nCallItemCount < pnCallMaxCount) Then
                Sheets(psModuleSheetName).OrderObject.SetQuote sItemCodes, pnSheetNum, pnCellY, nCol - nCallItemCount + pnCellX, _
                                                               pnDirection, pnTimer, pnTimerType
            Else
                Sheets(psModuleSheetName).OrderObject.SetQuote sItemCodes, pnSheetNum, pnCellY, nCol - pnCallMaxCount + pnCellX, _
                                                               pnDirection, pnTimer, pnTimerType
            End If
            
            '시세 호출 후 종목 리스트 초기화
            sItemCodes = ""
            nCallItemCount = 0

            'Delay 없는  연속 처리 금지
            WaitFor (pnWaitTime)
        End If
        
Continue:
    Next
    

End Sub


Public Sub StartQuote(pnMarketType As Integer, psItemSheetName As String, psModuleSheetName As String, _
                      pnItemY As Integer, pnItemX As Integer, _
                      pnTotalItemCount As Integer, pnCallMaxCount As Integer, _
                      Optional pnWaitTime As Single = 0.1)
'Parameters
'    pnMarketType As String       '시장 구분(1:장내종목(코스피, 코스닥), 2:선물옵션)
'    psItemSheetName As String    '종목 리스트가 있는 엑셀 시트 명
'    psModuleSheetName As String  '모듈이 있는 엑셀 시트 명
'    pnItemY As Integer           '세로 종목 위치
'    pnItemX As Integer           '가로 종목 위치
'    pnTotalItemCount As Integer  '총 종목수
'    pnCallMaxCount As Integer    '분할 호출 개수 : 최대 50종목까지만 호출 할 수 있음
'    pnWaitTime As Single         '시세 호출 후 대기 시간

    Dim nCol As Integer             '루프 변수
    Dim nCallItemCount As Integer   '호출한 종목 수
    Dim sItemCodes As String        '종목 리스트 ","로 구분한다.
    Dim sOrgCode As String          '주식 일때만 사용됨


    '여러번 호출 될 수 있기 때문에 여기에 있으면 안됨.
    'Application.Calculation = xlCalculationManual


    '종목 리스트 초기화
    sItemCodes = ""
    nCallItemCount = 0
    For nCol = 1 To pnTotalItemCount
        '종목 리스트 만들기
        '종목이 비어 있는지 확인
        If Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value <> "" Then
        '종목코드 값이 있는 경우
            If pnMarketType = 1 Then
            '주식(종목에 맨앞에"A"제거)
                sOrgCode = Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value
                sItemCodes = sItemCodes & Right(sOrgCode, Len(sOrgCode) - 1) & ","
            ElseIf pnMarketType = 2 Then
            '선물옵션
                sItemCodes = sItemCodes & Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value & ","
            End If
            
            nCallItemCount = nCallItemCount + 1
        Else
        '종목코드가 비어 있을 때
            '종목코드가 비어 있고 sItemCodes에 값이 존재 하는 경우
            If nCallItemCount > 0 Then
                '종목 리스트 마지막에 ","제거
                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
                
                '시세 호출
                Sheets(psModuleSheetName).OrderObject.StartQuote sItemCodes
                
                '시세 호출 후 종목 리스트 초기화
                sItemCodes = ""
                nCallItemCount = 0
    
                'Delay 없는  연속 처리 금지
                WaitFor (pnWaitTime)
                
            End If
            GoTo Continue
        
        End If

        
        '호출 할 종목 체크
        If (nCallItemCount > 0) And (((nCol Mod pnCallMaxCount) = 0) Or (nCol = pnTotalItemCount)) Then
            '종목 리스트 마지막에 ","제거
            sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
            
            '시세 호출
            Sheets(psModuleSheetName).OrderObject.StartQuote sItemCodes
            
            '시세 호출 후 종목 리스트 초기화
            sItemCodes = ""
            nCallItemCount = 0

            'Delay 없는  연속 처리 금지
            WaitFor (pnWaitTime)
        End If
        
Continue:
    Next


End Sub


Public Sub PauseQuote(pnMarketType As Integer, psItemSheetName As String, psModuleSheetName As String, _
                      pnItemY As Integer, pnItemX As Integer, _
                      pnTotalItemCount As Integer, pnCallMaxCount As Integer, _
                      Optional pnWaitTime As Single = 0.1)
'Parameters
'    pnMarketType As String       '시장 구분(1:장내종목(코스피, 코스닥), 2:선물옵션)
'    psItemSheetName As String    '종목 리스트가 있는 엑셀 시트 명
'    psModuleSheetName As String  '모듈이 있는 엑셀 시트 명
'    pnItemY As Integer           '세로 종목 위치
'    pnItemX As Integer           '가로 종목 위치
'    pnTotalItemCount As Integer  '총 종목수
'    pnCallMaxCount As Integer    '분할 호출 개수 : 최대 50종목까지만 호출 할 수 있음
'    pnWaitTime As Single         '시세 호출 후 대기 시간

    Dim nCol As Integer             '루프 변수
    Dim nCallItemCount As Integer   '호출한 종목 수
    Dim sItemCodes As String        '종목 리스트 ","로 구분한다.
    Dim sOrgCode As String          '주식 일때만 사용됨

    '여러번 호출 될 수 있기 때문에 여기에 있으면 안됨.
    'Application.Calculation = xlCalculationManual

    '종목 리스트 초기화
    sItemCodes = ""
    nCallItemCount = 0
    For nCol = 1 To pnTotalItemCount
        '종목 리스트 만들기
        '종목이 비어 있는지 확인
        If Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value <> "" Then
        '종목코드 값이 있는 경우
            If pnMarketType = 1 Then
            '주식(종목에 맨앞에"A"제거)
                sOrgCode = Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value
                sItemCodes = sItemCodes & Right(sOrgCode, Len(sOrgCode) - 1) & ","
            ElseIf pnMarketType = 2 Then
            '선물옵션
                sItemCodes = sItemCodes & Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value & ","
            End If
            
            nCallItemCount = nCallItemCount + 1
        Else
        '종목코드가 비어 있을 때
            '종목코드가 비어 있고 sItemCodes에 값이 존재 하는 경우
            If nCallItemCount > 0 Then
                '종목 리스트 마지막에 ","제거
                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
                
                '시세 호출
                Sheets(psModuleSheetName).OrderObject.StopQuote sItemCodes
                
                '시세 호출 후 종목 리스트 초기화
                sItemCodes = ""
                nCallItemCount = 0
    
                'Delay 없는  연속 처리 금지
                WaitFor (pnWaitTime)
                
            End If
            GoTo Continue
        
        End If

        
        '호출 할 종목 체크
        If (nCallItemCount > 0) And (((nCol Mod pnCallMaxCount) = 0) Or (nCol = pnTotalItemCount)) Then
            '종목 리스트 마지막에 ","제거
            sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
            
            '시세 호출
            Sheets(psModuleSheetName).OrderObject.StopQuote sItemCodes
            
            '시세 호출 후 종목 리스트 초기화
            sItemCodes = ""
            nCallItemCount = 0

            'Delay 없는  연속 처리 금지
            WaitFor (pnWaitTime)
        End If
        
Continue:
    Next
    
    '여러번 호출 될 수 있기 때문에 여기에 있으면 안됨.
    'Application.Calculation = xlCalculationAutomatic

End Sub


Public Sub StopQuote(pnMarketType As Integer, psItemSheetName As String, psModuleSheetName As String, _
                      pnItemY As Integer, pnItemX As Integer, _
                      pnTotalItemCount As Integer, pnCallMaxCount As Integer, _
                      Optional pnWaitTime As Single = 0.1)
'Parameters
'    pnMarketType As String       '시장 구분(1:장내종목(코스피, 코스닥), 2:선물옵션)
'    psItemSheetName As String    '종목 리스트가 있는 엑셀 시트 명
'    psModuleSheetName As String  '모듈이 있는 엑셀 시트 명
'    pnItemY As Integer           '세로 종목 위치
'    pnItemX As Integer           '가로 종목 위치
'    pnTotalItemCount As Integer  '총 종목수
'    pnCallMaxCount As Integer    '분할 호출 개수 : 최대 50종목까지만 호출 할 수 있음
'    pnWaitTime As Single         '시세 호출 후 대기 시간

    Dim nCol As Integer             '루프 변수
    Dim nCallItemCount As Integer   '호출한 종목 수
    Dim sItemCodes As String        '종목 리스트 ","로 구분한다.
    Dim sOrgCode As String          '주식 일때만 사용됨


    '종목 리스트 초기화
    sItemCodes = ""
    nCallItemCount = 0
    For nCol = 1 To pnTotalItemCount
        '종목 리스트 만들기
        '종목이 비어 있는지 확인
        If Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value <> "" Then
        '종목코드 값이 있는 경우
            If pnMarketType = 1 Then
            '주식(종목에 맨앞에"A"제거)
                sOrgCode = Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value
                sItemCodes = sItemCodes & Right(sOrgCode, Len(sOrgCode) - 1) & ","
            ElseIf pnMarketType = 2 Then
            '선물옵션
                sItemCodes = sItemCodes & Sheets(psItemSheetName).Cells(pnItemY, nCol + pnItemX - 1).Value & ","
            End If
            
            nCallItemCount = nCallItemCount + 1
        Else
        '종목코드가 비어 있을 때
            '종목코드가 비어 있고 sItemCodes에 값이 존재 하는 경우
            If nCallItemCount > 0 Then
                '종목 리스트 마지막에 ","제거
                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
                
                '시세 호출
                Sheets(psModuleSheetName).OrderObject.ResetQuote sItemCodes
                
                '시세 호출 후 종목 리스트 초기화
                sItemCodes = ""
                nCallItemCount = 0
    
                'Delay 없는  연속 처리 금지
                WaitFor (pnWaitTime)
                
            End If
            GoTo Continue
        
        End If

        
        '호출 할 종목 체크
        If (nCallItemCount > 0) And (((nCol Mod pnCallMaxCount) = 0) Or (nCol = pnTotalItemCount)) Then
            '종목 리스트 마지막에 ","제거
            sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
            
            '시세 호출
            Sheets(psModuleSheetName).OrderObject.ResetQuote sItemCodes
            
            '시세 호출 후 종목 리스트 초기화
            sItemCodes = ""
            nCallItemCount = 0

            'Delay 없는  연속 처리 금지
            WaitFor (pnWaitTime)
        End If
        
Continue:
    Next
    
    '여러번 호출 될 수 있어서 여기에 있으면 안됨.
    'Application.Calculation = xlCalculationAutomatic

End Sub



