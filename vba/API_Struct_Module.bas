Attribute VB_Name = "API_Struct_Module"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)




'grid의 갯수가 10개일수 밖에 없는 것 같음.
'http://support.microsoft.com/kb/179140/ko

Public Type TR1001_HOGA
    mdga As String * 9  '매도호가
    msga As String * 9  '매수호가
    dvol As String * 9  '매도호가수량
    svol As String * 9  '매수호가수량
    dcha As String * 9  '직전매도호가수량
    scha As String * 9  '직전매수호가수량
    dcnt As String * 9  '매도호가건수
    scnt As String * 9  '매수호가건수
End Type
 
Public Type TR1001_MOD
    cod2 As String * 12 'RTS 종목코드
    jmgb As String * 1  '종목구분(+KP, -KQ)
    curr As String * 9  '현재가
    diff As String * 9  '전일대비
    gvol As String * 12 '거래량
    gamt As String * 12 '거래대금
    jvol As String * 12 '전일거래량
    rate As String * 9  '등락률
    shga As String * 9  '상한가
    hhga As String * 9  '하한가
    gjga As String * 9  '기준가
    siga As String * 9  '시가
    koga As String * 9  '고가
    jega As String * 9  '저가
    jgbn As String * 3  '증거금율
    hoga(9) As TR1001_HOGA
    dvol As String * 9  '호가총수량 : 매수
    svol As String * 9  '호가총수량 : 매도
    dcha As String * 9  '직전대비총량 : 매도
    scha As String * 9  '직전대비총량 : 매수
    sdch As String * 9  '잔량차(svol-dvol)
    sum4 As String * 9  '종가합계 : 5일
    sum9 As String * 9  '종가합계 : 9일
    jggy As String * 9  '증거금율
    jqty As String * 9  '주문단위
End Type

Public Type TR1201_MOD
    nrec As String * 4  '반복횟수
    jmno As String * 6  '주문번호
    ojno As String * 6  '원주문번호
    emsg As String * 80 '오류메세지
End Type
 
Public Type TR1211_GRID
    juno As String * 10 '주문번호
    ojno As String * 10 '원주문번호
    cod2 As String * 12 '종목코드
    hnam As String * 40 '종목명
    odgb As String * 20 '주문구분
    mcgb As String * 20 '정취구분
    hogb As String * 20 '호가구분
    oprc As String * 12 '주문가격
    oqty As String * 12 '주문수량
    dprc As String * 12 '체결가격
    dqty As String * 12 '체결수량
    tqty As String * 12 '체결수량합
    wqty As String * 12 '미체결수량
    stat As String * 20 '접수상태
    time As String * 8  '주문시간
End Type
 
Public Type TR1211_MOD
    acno As String * 11 '계좌번호
    nrec As String * 4  '반복횟수
    grid(30) As TR1211_GRID
End Type

Public Type TR1221_GRID
    cod2 As String * 12 '종목코드
    hnam As String * 40 '종목명
    jgyh As String * 2  '잔고유형
    jqty As String * 10 '잔고수량
    xqty As String * 10 '매도가능
    pamt As String * 10 '매입평균가
    mamt As String * 15 '매입금액
    curr As String * 10 '현재가(*)
    rate As String * 10 '등락률
    diff As String * 10 '대비
    camt As String * 15 '평가금액
    tamt As String * 15 '평가손익(*)
    srat As String * 10 '평가수익률(*)
    sycd As String * 2  '신용코드
    sydt As String * 8  '대출일
    samt As String * 15 '신용금액
End Type

Public Type TR1221_MOD
    acno As String * 11 '계좌번호
    nrec As String * 4  '반복횟수
    grid(15) As TR1221_GRID
End Type


Public Type TR3001_HOGA
    mdga As String * 9  '매도호가
    msga As String * 9  '매수호가
    dvol As String * 9  '매도호가수량
    svol As String * 9  '매수호가수량
    dcnt As String * 9  '매도호가건수
    scnt As String * 9  '매수호가건수
End Type

Public Type TR3001_MOD
    curr As String * 9  '현재가
    diff As String * 9  '전일대비
    gvol As String * 12 '거래량
    gamt As String * 12 '거래대금
    rate As String * 9  '등락률
    shga As String * 9  '상한가
    hhga As String * 9  '하한가
    gjga As String * 9  '기준가
    siga As String * 9  '시가
    koga As String * 9  '고가
    jega As String * 9  '저가
    hoga(4) As TR3001_HOGA
    dvol As String * 9  '호가총수량 매도
    svol As String * 9  '호가총수량 매수
    dcnt As String * 9  '매도호가건수
    scnt As String * 9  '매수호가건수
    sdch As String * 9  '잔량차(svol -dvol)
    mgjv As String * 9  '미결제약정수량
End Type

Public Type TR3201_MOD
    nrec As String * 4  '반복횟수
    jmno As String * 6  '주문번호
    ojno As String * 6  '원주문번호
    emsg As String * 80 '오류메세지
End Type

Public Type TR3211_GRID
    mono As String * 6  '모주문번호
    juno As String * 6  '주문번호
    ojno As String * 6  '원주문번호
    cod2 As String * 8  '종목코드
    hnam As String * 30 '종목명
    odgb As String * 8  '주문구분
    hogb As String * 20 '주문유형
    oprc As String * 11 '주문가격
    oqty As String * 7  '주문수량
    dlgb As String * 4  '체결구분
    dprc As String * 11 '체결가격
    dqty As String * 7  '체결수량
    dtim As String * 6  '체결시간
    wqty As String * 7  '미체결수량
    hqty As String * 7  '정정/취소수량
    stat As String * 8  '처리상태
    time As String * 6  '처리시간
    jseq As String * 6  '접수번호
    yseq As String * 7  '약정번호
    ecod As String * 4  '거부코드
    dseq As String * 4  '체결횟수
End Type

Public Type TR3211_MOD
    acno As String * 11 '계좌번호
    nrec As String * 4  '반복횟수
    grid(30) As TR3211_GRID
End Type

Public Type TR3221_GRID
    cod2 As String * 8  '종목코드
    hnam As String * 30 '종목명
    dsgb As String * 6  '구분
    jqty As String * 10 '보유수량
    xqty As String * 10 '청산가능수량
    pamt As String * 10 '평균가/정산가
    curr As String * 10 '현재가
    diff As String * 10 '전일대비
    camt As String * 15 '평가금액
    tamt As String * 15 '평가손익
    srat As String * 10 '수익률
    mamt As String * 15 '매입금액
End Type

Public Type TR3221_MOD
    acno As String * 11 '계좌번호
    nrec As String * 4  '반복횟수
    grid(30) As TR3221_GRID
End Type

Public Type TR2001_GRID
    code As String * 10 ' RTS Symbol
    dvol As String * 12 ' 매도수량      333
    svol As String * 12 ' 매수수량      334
    rvol As String * 12 ' 순매수수량    343
    damt As String * 12 ' 매도금액      339
    samt As String * 12 ' 매수금액      340
    ramt As String * 12 ' 순매수금액    344
End Type

Public Type TR2001_MOD
    grid(12) As TR2001_GRID '외국인, 개인, 기관계, 투신, 금융투자, 보험, 은행, 기타금융, 연기금, 사모, 국가, 기타법인
End Type





