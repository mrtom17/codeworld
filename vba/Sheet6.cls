Attribute VB_Name = "Sheet6"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "AxIBK, 44, 0, IBKSCONNECTORLib, IBKSConnector"
Attribute VB_Control = "btnIBKSTR1211_1, 43, 1, MSForms, CommandButton"
Attribute VB_Control = "btnIBKSTR1211, 42, 2, MSForms, CommandButton"
Attribute VB_Control = "btnIBKSTR1221, 41, 3, MSForms, CommandButton"
Attribute VB_Control = "btnIBKSStockCancel, 40, 4, MSForms, CommandButton"
Attribute VB_Control = "btnIBKSStockSell, 39, 5, MSForms, CommandButton"
Attribute VB_Control = "btnIBKSStockBuy, 38, 6, MSForms, CommandButton"
Attribute VB_Control = "btnIBKSConnect, 37, 7, MSForms, CommandButton"
Attribute VB_Control = "btnIBKSInit, 36, 8, MSForms, CommandButton"
Attribute VB_Control = "btnSKReleaseReal, 35, 9, MSForms, CommandButton"
Attribute VB_Control = "btnSKRegistReal, 34, 10, MSForms, CommandButton"
Attribute VB_Control = "btnSKStock96013, 33, 11, MSForms, CommandButton"
Attribute VB_Control = "btnSKStock96007, 32, 12, MSForms, CommandButton"
Attribute VB_Control = "btnSKStock96012, 31, 13, MSForms, CommandButton"
Attribute VB_Control = "btnSKStockCancel, 30, 14, MSForms, CommandButton"
Attribute VB_Control = "btnSKStockSell, 29, 15, MSForms, CommandButton"
Attribute VB_Control = "btnSKStockBuy, 28, 16, MSForms, CommandButton"
Attribute VB_Control = "btnSKVersionExec, 26, 17, MSForms, CommandButton"
Attribute VB_Control = "btnSKAccountList, 25, 18, MSForms, CommandButton"
Attribute VB_Control = "btnLoginSK, 24, 19, MSForms, CommandButton"
Attribute VB_Control = "btnDisconnectSK, 23, 20, MSForms, CommandButton"
Attribute VB_Control = "btnConnectSK, 22, 21, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 21, 22, MSForms, CommandButton"
Attribute VB_Control = "OrderFirst, 20, 23, MSForms, CommandButton"
Attribute VB_Control = "OrderBatch, 19, 24, MSForms, CommandButton"
Attribute VB_Control = "CommandButton1, 18, 25, MSForms, CommandButton"
Attribute VB_Control = "btnFOHONBQ003, 17, 26, MSForms, CommandButton"
Attribute VB_Control = "btnStockHONAQ003, 16, 27, MSForms, CommandButton"
Attribute VB_Control = "btnFOHONBQ001, 15, 28, MSForms, CommandButton"
Attribute VB_Control = "btnStockHONAQ001, 14, 29, MSForms, CommandButton"
Attribute VB_Control = "FutureBatchCancel, 13, 30, MSForms, CommandButton"
Attribute VB_Control = "FutureBatchBuy, 12, 31, MSForms, CommandButton"
Attribute VB_Control = "FutureBatchSell, 11, 32, MSForms, CommandButton"
Attribute VB_Control = "StockBatchCancel, 10, 33, MSForms, CommandButton"
Attribute VB_Control = "StockBatchSell, 9, 34, MSForms, CommandButton"
Attribute VB_Control = "StockBatchBuy, 8, 35, MSForms, CommandButton"
Attribute VB_Control = "QuoteStart, 7, 36, MSForms, CommandButton"
Attribute VB_Control = "QuotePause, 6, 37, MSForms, CommandButton"
Attribute VB_Control = "QuoteStop, 5, 38, MSForms, CommandButton"
Attribute VB_Control = "QuoteCallFutures, 4, 39, MSForms, CommandButton"
Attribute VB_Control = "QuoteCallStock, 3, 40, MSForms, CommandButton"
Attribute VB_Control = "btnAccount, 2, 41, MSForms, CommandButton"
Attribute VB_Control = "btnLogin, 1, 42, MSForms, CommandButton"
Option Explicit
Enum TRKey
    TR1001 = 10     '주식시세조회
    TR1201          '주식주문
    TR1211          '주식체결조회
    TR1221          '주식잔고조회
    TR1222          '주식 실시간 잔고조회
    TR3001          '선옵시세조회
    TR3201          '선옵주문
    TR3211          '선옵체결조회
    TR3221          '선옵잔고조회
    TR3222          '선옵 실시간 잔고조회
    TR2001          '투자자별 매매동향
End Enum

Public WithEvents OrderObject As BizServer
Attribute OrderObject.VB_VarHelpID = -1

Public Sub AxIBK_OnAlert(ByVal code As Long, ByVal msg As String)
    Sheets("모듈").Cells(8, 9).Value = "[" & code & "]" & msg
'    MsgBox "[" & code & "]" & msg, vbInformation, "긴급전달메세지"
End Sub

Public Sub AxIBK_OnClose()
    Sheets("모듈").Cells(8, 9).Value = "접속이 종료되었습니다."
'    MsgBox "접속이 종료되었습니다.", vbInformation, "접속 종료"
End Sub

Private Sub AxIBK_OnContract(ByVal data As String)
    Dim handleValue As LongPtr
    
    'Sheets("모듈").Cells(44, 1).Value = data
    
    SendDataToMain GetHandleValue, MAIN_NOTICE, data

End Sub

Public Sub AxIBK_OnError(ByVal msg As String)
    Sheets("모듈").Cells(8, 9).Value = "[에러]" & msg
    'MsgBox "[에러]" & msg
End Sub

Public Sub AxIBK_OnGuideMsg(ByVal Key As Long, ByVal sMsg As String)
'    MsgBox "[" & key & "]" & sMsg, vbOKOnly, "코스콤"
End Sub

Public Sub AxIBK_OnLogin(ByVal bLogin As Boolean, ByVal msg As String)
    Sheets("모듈").Cells(8, 9).Value = "[" & bLogin & "] " & msg
'    MsgBox "[" & bLogin & "] " & msg
End Sub

Public Sub AxIBK_OnRecvData(ByVal Key As Long, ByVal dptr As Long, ByVal size As Long, ByVal bNext As Boolean, ByVal nkey As String)
    Dim str() As Byte
    Dim dststr() As Byte
    Dim strData As String
    Dim ii As Integer
    Dim strTemp As String
    Dim nCount As Long
    
    Dim tempRowStr As String
    Dim finalStr As String
    
    Dim nCount1 As Integer
    Dim strCount1 As String
    Dim strKeyCount1 As String
    Dim strKeyName1 As String
    Dim i, iRow, iCol, c, pos As Integer
    
    
    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)

'    ReDim str(size)
'    CopyMemory str(0), ByVal dptr + 15, size - 15
'    strData = StrConv(str, vbUnicode)
'    MsgBox strData, vbOKOnly, key
    
    Select Case Key
'        Case TRKey.TR1001 ' 주식현재가 조회
'            Dim TR1001 As TR1001_MOD
'
'            CopyMemory TR1001, ByVal dptr, size
'
'            'TR1001구조체에 접근하여 값을 셀에 표시하면 됩니다.
        
        Case TRKey.TR1201 ' 주식주문
            Dim TR1201 As TR1201_MOD
            
            CopyMemory TR1201, ByVal dptr, size
'            nrec As String * 4  '반복횟수
'            jmno As String * 6  '주문번호
'            ojno As String * 6  '원주문번호
'            emsg As String * 80 '오류메세지
            
            'TR1201구조체에 접근하여 값을 셀에 표시하면 됩니다.
            Sheets("모듈").Cells(8, 9).Value = "주문번호" & TR1201.jmno & " 메세지 " & TR1201.emsg
            
         Case TRKey.TR1211 '주식체결내역조회
            Dim TR1211 As TR1211_MOD
            
            CopyMemory TR1211, ByVal dptr, size
'            acno As String * 11 '계좌번호
'            nrec As String * 4  '반복횟수
'            grid(10) As TR1211_GRID

'            juno As String * 5  '주문번호
'            ojno As String * 5  '원주문번호
'            cod2 As String * 12 '종목코드
'            hnam As String * 40 '종목명
'            odgb As String * 20 '주문구분
'            mcgb As String * 20 '정취구분--
'            hogb As String * 20 '호가구분--
'            oprc As String * 12 '주문가격
'            oqty As String * 12 '주문수량
'            dprc As String * 12 '체결가격--
'            dqty As String * 12 '체결수량--
'            tqty As String * 12 '체결수량합--
'            wqty As String * 12 '미체결수량
'            stat As String * 20 '접수상태
'            time As String * 8  '주문시간

            'TR1211구조체에 접근하여 값을 셀에 표시하면 됩니다.
             
            
            If HONAQ003Count_T = 0 Then
                HONAQ003Count_T = Int(TR1211.nrec)
                If Int(TR1211.nrec) <> 0 Then
                    'ReDim HONAQ003Data(HONAQ003Count_T, 14) As String
                    ReDim HONAQ003Data(1000, 14) As String  ' 임시 방편....
                End If
            Else
                HONAQ003Count_T = HONAQ003Count_T + Int(TR1211.nrec)
                'ReDim Preserve HONAQ003Data(HONAQ003Count_T, 14) As String
            End If

            If isSelectHONAQ003 Then
                Application.Calculation = xlCalculationManual
            End If
        
            HONAQ003Count = HONAQ003Count_T
            
            '배열에 넣어줌.
            
            If Int(TR1211.nrec) <> 0 Then
                'ReDim TR1211.grid(TR1211.nrec) As TR1211_GRID
'                For i = 0 To Int(TR1211.nrec) - 1
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 0) = Trim(TR1211.grid(i).hnam) '종목명
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 1) = Trim(TR1211.grid(i).odgb) '주문구분
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 2) = Trim(TR1211.grid(i).oqty) '주문수량
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 3) = Trim(TR1211.grid(i).dqty) '체결수량
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 4) = Trim(TR1211.grid(i).wqty) '미체결수량
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 5) = Trim(TR1211.grid(i).oprc) '주문가격
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 6) = Trim(TR1211.grid(i).dprc) '체결가격
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 7) = Trim(TR1211.grid(i).ojno) '주문번호
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 8) = Trim(TR1211.grid(i).juno) '원주문번호
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 9) = Trim("")
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 10) = "A" + Trim(TR1211.grid(i).cod2) '종목코드
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 11) = Trim("")
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 12) = Trim(TR1211.grid(i).stat)
'                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 13) = Trim(TR1211.grid(i).time)
'                Next i


'                juno As String * 10 '주문번호1, 10, Trim(StrConv(MidB(str, 1, 10), vbUnicode))
'                ojno As String * 10 '원주문번호11, 10, Trim(StrConv(MidB(str, 11, 10), vbUnicode))
'                cod2 As String * 12 '종목코드21, 12, Trim(StrConv(MidB(str, 21, 12), vbUnicode))
'                hnam As String * 40 '종목명33, 40, Trim(StrConv(MidB(str, 33, 40), vbUnicode))
'                odgb As String * 20 '주문구분73, 20, Trim(StrConv(MidB(str, 73, 20), vbUnicode))
'                mcgb As String * 20 '매매구분93, 20, Trim(StrConv(MidB(str, 93, 20), vbUnicode))
'                hogb As String * 20 '호가구분113, 20, Trim(StrConv(MidB(str, 113, 20), vbUnicode))
'                oprc As String * 12 '주문가격133, 12, Trim(StrConv(MidB(str, 133, 12), vbUnicode))
'                oqty As String * 12 '주문수량145, 12, Trim(StrConv(MidB(str, 145, 12), vbUnicode))
'                dprc As String * 12 '체결가격157, 12, Trim(StrConv(MidB(str, 157, 12), vbUnicode))
'                dqty As String * 12 '체결수량169, 12, Trim(StrConv(MidB(str, 169, 12), vbUnicode))
'                tqty As String * 12 '체결수량합181, 12, Trim(StrConv(MidB(str, 181, 12), vbUnicode))
'                wqty As String * 12 '미체결수량193, 12, Trim(StrConv(MidB(str, 193, 12), vbUnicode))
'                stat As String * 20 '접수상태205, 20, Trim(StrConv(MidB(str, 205, 20), vbUnicode))
'                time As String * 8  '주문시간225, 8, Trim(StrConv(MidB(str, 225, 8), vbUnicode))


                               
                For i = 0 To Int(TR1211.nrec) - 1
                    ReDim Preserve str(232)
                    pos = 15 + (i * 232)
                    CopyMemory str(0), ByVal dptr + pos, 232
                    
                    ' 각 필드에서 Trim(StrConv(MidB(변환), vbUnicode)) 패턴을 적용합니다.
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 0) = Trim(StrConv(MidB(str, 33, 40), vbUnicode)) '종목명
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 1) = Trim(StrConv(MidB(str, 73, 20), vbUnicode)) '주문구분
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 2) = Trim(StrConv(MidB(str, 145, 12), vbUnicode)) '주문수량
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 3) = Trim(StrConv(MidB(str, 169, 12), vbUnicode)) '체결수량
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 4) = Trim(StrConv(MidB(str, 193, 12), vbUnicode)) '미체결수량
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 5) = Trim(StrConv(MidB(str, 133, 12), vbUnicode)) '주문가격
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 6) = Trim(StrConv(MidB(str, 157, 12), vbUnicode)) '체결가격
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 7) = Trim(StrConv(MidB(str, 11, 10), vbUnicode)) '원주문번호
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 8) = Trim(StrConv(MidB(str, 1, 10), vbUnicode)) '주문번호
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 9) = Trim("") ' 빈 필드
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 10) = "A" + Trim(StrConv(MidB(str, 21, 12), vbUnicode)) '종목코드
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 11) = Trim("") ' 빈 필드
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 12) = Trim(StrConv(MidB(str, 205, 20), vbUnicode)) '접수상태
                    HONAQ003Data(HONAQ003Count_T - Int(TR1211.nrec) + i, 13) = Trim(StrConv(MidB(str, 225, 8), vbUnicode)) '주문시간
                Next i


            
                ' 연속된 데이터가 있으면 다시 미체결조회
                If bNext Then
                    If AxIBK.TR1211(TRKey.TR1211, Sheets("모듈").Cells(6, 12).Value, Sheets("모듈").Cells(7, 12).Value, "", 0, 0, 2, 2, nkey) = False Then
                        MsgBox AxIBK.GetLastErrMsg
                    End If
                    Exit Sub
                Else
                    HONAQ003Count = HONAQ003Count_T
                End If
                 
    
                
                ' 화면에 출력함
                ' 화면 출력 안함.
                'isSelectHONAQ003 = True
                If isSelectHONAQ003 Then
                    For iRow = 0 To HONAQ003Count - 1
                        For iCol = 0 To UBound(HONAQ003Data, 2)
                            strData = HONAQ003Data(iRow, iCol)
                            Select Case iCol
                            Case 7 '주문번호
                                Sheets("체결내역").Cells(iRow + 3, 1).Value = strData
                            Case 8 '원주문번호
                                Sheets("체결내역").Cells(iRow + 3, 2).Value = strData
                            Case 13 '주문시간
                                Sheets("체결내역").Cells(iRow + 3, 3).Value = strData
                            Case 10 '종목코드
                                Sheets("체결내역").Cells(iRow + 3, 4).Value = strData
                            Case 0 '종목명
                                Sheets("체결내역").Cells(iRow + 3, 5).Value = strData
                            Case 1 '주문구분명
                                Sheets("체결내역").Cells(iRow + 3, 6).Value = strData
                            Case 2 '주문수량
                                Sheets("체결내역").Cells(iRow + 3, 7).Value = strData
                            Case 5 '주문가
                                Sheets("체결내역").Cells(iRow + 3, 8).Value = strData
                            Case 3 '체결수량
                                Sheets("체결내역").Cells(iRow + 3, 9).Value = strData
                            Case 6 '체결가
                                Sheets("체결내역").Cells(iRow + 3, 10).Value = strData
                            Case 4 '미체결수량
                                Sheets("체결내역").Cells(iRow + 3, 11).Value = strData
                            Case 12 '체결상태명
                                Sheets("체결내역").Cells(iRow + 3, 12).Value = strData
                            End Select
                        Next
                    Next
                End If
                
                'main으로 데이터 전달
                For iRow = 0 To HONAQ003Count - 1
                    tempRowStr = "" ' 각 row에 대한 임시 문자열 초기화
                    
                    For iCol = 0 To UBound(HONAQ003Data, 2)
                        If iCol = UBound(HONAQ003Data, 2) Then
                            tempRowStr = tempRowStr & Trim(HONAQ003Data(iRow, iCol)) ' 마지막 열이면 탭 추가 안 함
                        Else
                            tempRowStr = tempRowStr & Trim(HONAQ003Data(iRow, iCol)) & vbTab ' 데이터 사이에 탭 추가
                        End If
                    Next iCol
                    
                    If iRow = HONAQ003Count - 1 Then
                        finalStr = finalStr & tempRowStr ' 마지막 row이면 row 구분자 추가 안 함
                    Else
                        finalStr = finalStr & tempRowStr & "|" ' row 데이터 사이에 | 추가
                    End If
                Next iRow
                
                SendDataToMain GetHandleValue, MAIN_BUSI_RECV_HONAQ003, finalStr

                If isSelectHONAQ003 Then
                    Application.Calculation = xlCalculationAutomatic
                End If
            End If

'         Case TRKey.TR3001 '선옵현재가 조회
'            Dim TR3001 As TR3001_MOD
'
'            CopyMemory TR3001, ByVal dptr, size
'
'            'TR3001구조체에 접근하여 값을 셀에 표시하면 됩니다.
'
'         Case TRKey.TR3201 '선옵주문
'            Dim TR3201 As TR3201_MOD
'
'            CopyMemory TR3201, ByVal dptr, size
'
'            'TR3201구조체에 접근하여 값을 셀에 표시하면 됩니다.
'
'         Case TRKey.TR3211 '선옵체결내역 조회
'            Dim TR3211 As TR3211_MOD
'
'            CopyMemory TR3211, ByVal dptr, size
'
'            'TR3211구조체에 접근하여 값을 셀에 표시하면 됩니다.
'
'         Case TRKey.TR3221 '선옵잔고내역 조회
'            Dim TR3221 As TR3221_MOD
'
'            CopyMemory TR3221, ByVal dptr, size
'
'            'TR3221구조체에 접근하여 값을 셀에 표시하면 됩니다.
        
        Case TRKey.TR1221 '주식잔고내역 조회
            'NextVal = TR1221.nrec
'            acno As String * 11 '계좌번호
'            nrec As String * 4  '반복횟수
'            grid(10) As TR1221_GRID
        
'            cod2 As String * 12 '종목코드
'            hnam As String * 40 '종목명
'            jgyh As String * 2  '잔고유형 --
'            jqty As String * 10 '잔고수량
'            xqty As String * 10 '매도가능
'            pamt As String * 10 '매입평균가--
'            mamt As String * 15 '매입금액
'            curr As String * 10 '현재가(*)
'            rate As String * 10 '등락률
'            diff As String * 10 '대비
'            camt As String * 15 '평가금액
'            tamt As String * 15 '평가손익(*)
'            srat As String * 10 '평가수익률(*)
'            sycd As String * 2  '신용코드
'            sydt As String * 8  '대출일
'            samt As String * 15 '신용금액--

            Dim TR1221 As TR1221_MOD
'            Dim TR1221_Data()
'            Dim k As Integer

            CopyMemory TR1221, ByVal dptr, size
'            ReDim TR1221_Data(Int(TR1221.nrec))
'            CopyMemory TR1221_Data(0), ByVal dptr + 15, size - 15
'
'            If (size > 0) Then
'                nCount = (size - 15) / 194
'                ReDim TR1221_Data(nLen)
'            End If

            
            If HONAQ001Count_T = 0 Then
                HONAQ001Count_T = Int(TR1221.nrec)
                If Int(TR1221.nrec) <> 0 Then
                    'ReDim HONAQ001Data(HONAQ001Count_T, 21) As String
                    ReDim HONAQ001Data(1000, 21)
                    'ReDim Preserve TR1221.grid(Int(TR1221.nrec)) As TR1221_GRID
                End If
            Else
                HONAQ001Count_T = HONAQ001Count_T + Int(TR1221.nrec)
                'ReDim Preserve HONAQ001Data(HONAQ001Count_T, 21) As String
            End If


            Application.Calculation = xlCalculationManual
            
            For c = QuoteStockHStartPos To QuoteStockItemCount + QuoteStockHStartPos - 1
                Sheets("가격수령").Cells(15, c).Value = 0
                Sheets("가격수령").Cells(16, c).Value = 0
            Next
            Application.Calculation = xlCalculationAutomatic

            HONAQ001Count = HONAQ001Count_T

            '배열에 넣어줌.
'            CurrPos = 0
'            If (CurrPos = 0) Then
'                CurrPos = 15 + (194 * 10)
'            Else
'                CurrPos = CurrVal + (194 * 10)
'            End If
'
'
'            If (TR1221.nrec <= 10) Then
'                CopyMemory TR1221, ByVal dptr, size
'            Else
'
'            End If
'
'            CurrPos = 0
'            For k = 0 To nCount - 1
'                If (k = 0) Then
'                    CurrPos = (k * 10)
'                End If
'                For i = 0 To CurrPos - 1
            If Int(TR1221.nrec) <> 0 Then
                'ReDim TR1211.grid(TR1211.nrec) As TR1211_GRID
'                For i = 0 To Int(TR1221.nrec) - 1
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 0) = Trim(TR1221.grid(i).sycd) '신용금액
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 1) = Trim("")
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 2) = Trim(TR1221.grid(i).hnam) '종목명
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 3) = Trim(TR1221.grid(i).xqty) '매도가능
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 4) = Trim(TR1221.grid(i).jqty) '잔고수량
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 5) = Trim("")
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 6) = Trim(TR1221.grid(i).rate) '등락률
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 7) = Trim(TR1221.grid(i).srat) '평가수익률(*)
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 8) = Trim(TR1221.grid(i).tamt) '평가손익(*)
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 9) = Trim(TR1221.grid(i).mamt) '매입금액
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 10) = Trim(TR1221.grid(i).camt) '평가금액
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 11) = Trim("")
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 12) = Trim(TR1221.grid(i).sydt) '대출일
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 13) = Trim("")
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 14) = Trim(TR1221.grid(i).curr) '현재가(*)
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 15) = Trim("")
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 16) = Trim("")
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 17) = Trim("")
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 18) = Trim(TR1221.grid(i).cod2) '종목코드
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 19) = Trim(TR1221.grid(i).diff) '대비
'                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 20) = Trim("")
'                Next i
    
'                cod2 As String * 12 '종목코드1, 12, Trim(Mid(strData, 1, 12))
'                hnam As String * 40 '종목명13, 40, Trim(Mid(strData, 13, 40))
'                jgyh As String * 2  '잔고유형53, 2, Trim(Mid(strData, 53, 2))
'                jqty As String * 10 '잔고수량55, 10, Trim(Mid(strData, 55, 10))
'                xqty As String * 10 '매도가능65, 10, Trim(Mid(strData, 65, 10))
'                pamt As String * 10 '매입평균가75, 10, Trim(Mid(strData, 75, 10))
'                mamt As String * 15 '매입금액85, 15, Trim(Mid(strData, 85, 15))
'                curr As String * 10 '현재가(*)100, 10, Trim(Mid(strData, 100, 10))
'                rate As String * 10 '등락률110, 10, Trim(Mid(strData, 110, 10))
'                diff As String * 10 '대비120, 10, Trim(Mid(strData, 120, 10))
'                camt As String * 15 '평가금액130, 15, Trim(Mid(strData, 130, 15))
'                tamt As String * 15 '평가손익(*)145, 15, Trim(Mid(strData, 145, 15))
'                srat As String * 10 '평가수익률(*)160, 10, Trim(Mid(strData, 160, 10))
'                sycd As String * 2  '신용코드170, 2, Trim(Mid(strData, 170, 2))
'                sydt As String * 8  '대출일172, 8, Trim(Mid(strData, 172, 8))
'                samt As String * 15 '신용금액180, 15, Trim(Mid(strData, 180, 15))
                
                For i = 0 To Int(TR1221.nrec) - 1
    
                    ReDim Preserve str(194)
                    pos = 15 + (i * 194)
                    CopyMemory str(0), ByVal dptr + pos, 194
                    'strData = StrConv(s, vbUnicode)
                    'MsgBox Mid(strData, 1, 12), vbOKOnly, i + 1
                    'MsgBox StrConv(MidB(str, 13, 40), vbUnicode), vbOKOnly, i + 1
                    
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 0) = Trim(StrConv(MidB(str, 180, 15), vbUnicode)) '신용금액
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 1) = Trim("")
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 2) = Trim(StrConv(MidB(str, 13, 40), vbUnicode)) '종목명
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 3) = Trim(StrConv(MidB(str, 65, 10), vbUnicode)) '매도가능
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 4) = Trim(StrConv(MidB(str, 55, 10), vbUnicode)) '잔고수량
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 5) = Trim("")
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 6) = Trim(StrConv(MidB(str, 110, 10), vbUnicode)) '등락률
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 7) = Trim(StrConv(MidB(str, 160, 10), vbUnicode)) '평가수익률(*)
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 8) = Trim(StrConv(MidB(str, 145, 15), vbUnicode)) '평가손익(*)
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 9) = Trim(StrConv(MidB(str, 85, 15), vbUnicode)) '매입금액
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 10) = Trim(StrConv(MidB(str, 130, 15), vbUnicode)) '평가금액
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 11) = Trim("")
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 12) = Trim(StrConv(MidB(str, 172, 8), vbUnicode)) '대출일
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 13) = Trim("")
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 14) = Trim(StrConv(MidB(str, 100, 10), vbUnicode)) '현재가(*)
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 15) = Trim("")
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 16) = Trim("")
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 17) = Trim("")
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 18) = Trim(StrConv(MidB(str, 1, 12), vbUnicode)) '종목코드
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 19) = Trim(StrConv(MidB(str, 120, 10), vbUnicode)) '대비
                    HONAQ001Data(HONAQ001Count_T - Int(TR1221.nrec) + i, 20) = Trim("")

                Next i
                
                
                ' 연속된 데이터가 있으면 다시 잔고조회
                If bNext Then
                    If AxIBK.TR1221(TRKey.TR1221, Sheets("모듈").Cells(6, 12).Value, Sheets("모듈").Cells(7, 12).Value, 1, nkey) = False Then
                        MsgBox AxIBK.GetLastErrMsg
                    End If
                    Exit Sub
                Else
                    HONAQ001Count = HONAQ001Count_T
                End If
                 
    
                
                ' 화면에 출력함
                Application.Calculation = xlCalculationManual
                
'                For c = QuoteStockHStartPos To QuoteStockItemCount + QuoteStockHStartPos - 1
'                    Sheets("가격수령").Cells(15, c).Value = 0
'                    Sheets("가격수령").Cells(16, c).Value = 0
'                Next
                
                For iRow = 0 To HONAQ001Count - 1
                    For c = QuoteStockHStartPos To QuoteStockItemCount + QuoteStockHStartPos - 1
                        If (Sheets("가격수령").Cells(3, c).Value = HONAQ001Data(iRow, 18)) Then
                            Sheets("가격수령").Cells(15, c).Value = HONAQ001Data(iRow, 3)
                            Sheets("가격수령").Cells(16, c).Value = CStr(CLng(HONAQ001Data(iRow, 9)) / CLng(HONAQ001Data(iRow, 3)))
                        End If
                    Next
                Next
                
                'main으로 데이터 전달
                For iRow = 0 To HONAQ001Count - 1
                    tempRowStr = "" ' 각 row에 대한 임시 문자열 초기화
                    
                    For iCol = 0 To UBound(HONAQ001Data, 2)
                        If iCol = UBound(HONAQ001Data, 2) Then
                            tempRowStr = tempRowStr & Trim(HONAQ001Data(iRow, iCol)) ' 마지막 열이면 탭 추가 안 함
                        Else
                            tempRowStr = tempRowStr & Trim(HONAQ001Data(iRow, iCol)) & vbTab ' 데이터 사이에 탭 추가
                        End If
                    Next iCol
                    
                    If iRow = HONAQ001Count - 1 Then
                        finalStr = finalStr & tempRowStr ' 마지막 row이면 row 구분자 추가 안 함
                    Else
                        finalStr = finalStr & tempRowStr & "|" ' row 데이터 사이에 | 추가
                    End If
                Next iRow
                
                SendDataToMain GetHandleValue, MAIN_BUSI_RECV_HONAQ001, finalStr

                Application.Calculation = xlCalculationAutomatic
            End If
'            Next
    End Select

End Sub

Private Sub AxIBK_OnRecvDataID(ByVal Key As Long, ByVal dptr As Long, ByVal size As Long, ByVal bNext As Boolean, ByVal nkey As String, ByVal id As Long)

End Sub

Public Sub AxIBK_OnUpdateEnd()
    MsgBox "파일업데이트 완료", vbInformation, "파일 업데이트"
End Sub

Public Sub AxIBK_OnUpdateStart()
    MsgBox "파일업데이트 시작", vbInformation, "파일 업데이트"
End Sub

Public Sub AxIBK_OnVerUpdate()
    MsgBox "플랫폼 업데이트 완료" & vbCrLf & "파일복사 후 재기동 하시기 바랍니다.", vbInformation, "플랫폼파일 업데이트"
End Sub

Public Sub btnAccount_Click()
    Sheets("모듈").OrderObject.GetAccountList
End Sub

Public Sub btnConnectSK_Click()
' 서버에 접속한다

'    Dim strID As String
'    Dim strPassword As String
'    Dim strCertPassword As String
'
'    strID = Sheets("모듈").Cells(1, 9).Value
'    strPassword = Sheets("모듈").Cells(2, 9).Value
'    strCertPassword = Sheets("모듈").Cells(3, 9).Value
'
'    WapiConnectMngCtrl.id = strID
'    WapiConnectMngCtrl.IDPassword = strPassword
'    WapiConnectMngCtrl.CertPassword = strCertPassword
'    WapiConnectMngCtrl.SaveCertPass = True
'    WapiConnectMngCtrl.KeyboardSecurity = False
'    WapiConnectMngCtrl.FireWall = False
'    WapiConnectMngCtrl.Proxy = False
'
'
'    Dim lResult As Long
'    Dim bLoginDlg As Boolean
'
'    bLoginDlg = False
'
'    lResult = WapiConnectMngCtrl.ConnectWapi(bLoginDlg)
'
'
'
'    'If lResult = 1 Then
'    If lResult = 0 Or lResult = 100 Then
'        'MsgBox "서버에 접속했습니다."
'        Sheets("모듈").Cells(8, 9).Value = "서버에 접속했습니다."
'    Else
'        'MsgBox "서버에 접속하지 못했습니다."
'        Sheets("모듈").Cells(8, 9).Value = "서버에 접속하지 못했습니다."
'    End If
            
End Sub

Public Sub btnDisconnectSK_Click()
    ' 서버와 연결을 해제한다
'    Dim lResult As Long
'
'    lResult = WapiConnectMngCtrl.DisconnectWapi()
'
'    If lResult = 1 Then
'        Sheets("모듈").Cells(8, 9).Value = "서버에 연결이 종료 되었습니다."
'    Else
'        Sheets("모듈").Cells(8, 9).Value = "서버에 연결이 종료가 실패 되었습니다."
'    End If
    
End Sub

Public Sub btnFOHONBQ001_Click()

    Dim sIssueCode As String     '계좌번호
    Dim sPassword As String     '비밀번호
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value
    
    sIssueCode = AccountListFO(0, 0)
    sPassword = AccountPass
    
    CreateBizObject
    
    HONBQ001Count = -1
    
    Sheets("모듈").OrderObject.SendReceiveEx sIssueCode, sPassword, "HONBQ001"   '선물옵션잔고조회

End Sub

Public Sub btnFOHONBQ003_Click()
    Dim sIssueCode As String     '계좌번호
    Dim sPassword As String     '비밀번호
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value
    
    sIssueCode = AccountListFO(0, 0)
    sPassword = AccountPass
    
    CreateBizObject
    
    HONBQ003Count = -1
    
    Sheets("모듈").OrderObject.SendReceiveEx sIssueCode, sPassword, "HONBQ003", "2"   '체결내역조회 조회구분 "0":전체(체결,미체결), "1":체결, "2":미체결
End Sub

Public Sub btnIBKSConnect_Click()
    Dim usid, uspw, uscpw, svrip, svrpt
    usid = Sheets("모듈").Cells(3, 12).Value
    uspw = Sheets("모듈").Cells(4, 12).Value
    uscpw = Sheets("모듈").Cells(5, 12).Value
    svrip = Sheets("모듈").Cells(1, 12).Value
    svrpt = Sheets("모듈").Cells(2, 12).Value
    AxIBK.Login usid, uspw, uscpw, svrip, svrpt

End Sub

Public Sub btnIBKSInit_Click()
    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)

    If AxIBK.Initialize() Then
        'MsgBox "초기화 성공"
    Else
        MsgBox AxIBK.GetLastErrMsg()
    End If

End Sub

Public Sub btnIBKSStockBuy_Click()
    Dim acno, pswd, mmgb, ojno, code, jqty, jprc, hogb, mdgb
    Dim i As Integer

    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)
    
    mmgb = 2    'int     매도(1), 매수(2), 정정(3), 취소(4)
    acno = Sheets("모듈").Cells(6, 12).Value    'char    11  계좌번호
    pswd = Sheets("모듈").Cells(7, 12).Value    'char    8   계좌비밀번호
    ojno = 0    'int     원주문번호(정정/취소시)
    hogb = 0    'int     "주문구분'일반(0), 시장가(3), 조건부지정가(5),'최유리지정가 (6), 최우선지정가(7)'-> IOC주문일 경우 + 10 , FOK주문일 경우 + 20'-> 주문가능호가는 HTS참조"
    mdgb = 0    'int     정정취소구분: 일부(1), 전부(2)


    '주식매수주문
    For i = 1 To QuoteStockItemCount
        code = Sheets("가격수령").Cells(3, i + QuoteStockHStartPos - 1).Value    'char    7   "주식단축코드(7자리)'일반주식:  'A' + 단축코드'ELW:      'J' + 단축코드"
        jprc = CLng(Sheets("가격수령").Cells(29, i + QuoteStockHStartPos - 1).Value)    'int     주문가격
        jqty = CLng(Sheets("가격수령").Cells(30, i + QuoteStockHStartPos - 1).Value)    'int     주문수량
        
        If (jprc <> 0) And (jqty <> 0) Then

            If AxIBK.TR1201(TRKey.TR1201, mmgb, acno, pswd, ojno, code, jqty, jprc, hogb, mdgb) = False Then
            'If AxIBK.TR1201(TRKey.TR1201, "2", sAccount, AccountPass, "0", sIssueCode, sQty, sPrice, sCallType, 0) = False Then
                'MsgBox AxIBK.GetLastErrMsg
            End If

            WaitFor (0.1)
        End If
                
    Next
        
End Sub

Public Sub btnIBKSStockCancel_Click()
    Dim acno, pswd, mmgb, ojno, code, jqty, jprc, hogb, mdgb
    Dim iRow As Integer
     
    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)
    
    mmgb = 4    'int     매도(1), 매수(2), 정정(3), 취소(4)
    acno = Sheets("모듈").Cells(6, 12).Value    'char    11  계좌번호
    pswd = Sheets("모듈").Cells(7, 12).Value    'char    8   계좌비밀번호
    jprc = 0    'int     주문가격
    hogb = 0    'int     "주문구분'일반(0), 시장가(3), 조건부지정가(5),'최유리지정가 (6), 최우선지정가(7)'-> IOC주문일 경우 + 10 , FOK주문일 경우 + 20'-> 주문가능호가는 HTS참조"
    mdgb = 2    'int     정정취소구분: 일부(1), 전부(2)
    
    
    
    '주식주문취소
    For iRow = 0 To HONAQ003Count - 1
        
        code = HONAQ003Data(iRow, 10)    'char    7   "주식단축코드(7자리)'일반주식:  'A' + 단축코드'ELW:      'J' + 단축코드"
        If HONAQ003Data(iRow, 4) = "" Then
            jqty = 0
        Else
            jqty = CLng(CInt(HONAQ003Data(iRow, 4)))    'int     주문수량
        End If
        
        
        If HONAQ003Data(iRow, 8) = "" Then
            ojno = 0
        Else
            ojno = CLng(HONAQ003Data(iRow, 8))    'int     원주문번호(정정/취소시)
        End If

        If AxIBK.TR1201(TRKey.TR1201, mmgb, acno, pswd, ojno, code, jqty, jprc, hogb, mdgb) = False Then
        'If AxIBK.TR1201(TRKey.TR1201, "4", sAccount, AccountPass, sOrgOrdNum, sIssueCode, sQty, "0", sCallType, 0) = False Then
            'MsgBox AxIBK.GetLastErrMsg
        End If
        WaitFor (0.1)
    Next
   
End Sub

Public Sub btnIBKSStockSell_Click()
    Dim acno, pswd, mmgb, ojno, code, jqty, jprc, hogb, mdgb
    Dim i As Integer
    
    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)
    
    mmgb = 1    'int     매도(1), 매수(2), 정정(3), 취소(4)
    acno = Sheets("모듈").Cells(6, 12).Value    'char    11  계좌번호
    pswd = Sheets("모듈").Cells(7, 12).Value    'char    8   계좌비밀번호
    ojno = 0    'int     원주문번호(정정/취소시)
    hogb = 0    'int     "주문구분'일반(0), 시장가(3), 조건부지정가(5),'최유리지정가 (6), 최우선지정가(7)'-> IOC주문일 경우 + 10 , FOK주문일 경우 + 20'-> 주문가능호가는 HTS참조"
    mdgb = 0    'int     정정취소구분: 일부(1), 전부(2)
    
    '주식매도주문
        
    For i = 1 To QuoteStockItemCount
        code = Sheets("가격수령").Cells(3, i + QuoteStockHStartPos - 1).Value    'char    7   "주식단축코드(7자리)'일반주식:  'A' + 단축코드'ELW:      'J' + 단축코드"
        jprc = CLng(Sheets("가격수령").Cells(31, i + QuoteStockHStartPos - 1).Value)    'int     주문가격
        jqty = CLng(Sheets("가격수령").Cells(32, i + QuoteStockHStartPos - 1).Value)    'int     주문수량
        
        If (jprc <> 0) And (jqty <> 0) Then
        
            If AxIBK.TR1201(TRKey.TR1201, mmgb, acno, pswd, ojno, code, jqty, jprc, hogb, mdgb) = False Then
            'If AxIBK.TR1201(TRKey.TR1201, "1", sAccount, AccountPass, "0", sIssueCode, sQty, sPrice, sCallType, 0) = False Then
                'MsgBox AxIBK.GetLastErrMsg
            End If
            WaitFor (0.1)
        End If
                
    Next

End Sub

Public Sub btnIBKSTR1211_1_Click()
    '주식:주식 체결내역 조회-미체결(TR1211)
    Dim Key, acno, pswd, code, dsgb, sygb, dlgb, sort, nkey

    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)
    
    HONAQ003Count = -1
    HONAQ003Count_T = 0
    isSelectHONAQ003 = False
    
    acno = Sheets("모듈").Cells(6, 12).Value    '    char    11  계좌번호
    pswd = Sheets("모듈").Cells(7, 12).Value    '    char    8   계좌비밀번호
    code = ""   '    char    7   [Option]조회종목번호
    dsgb = 0    '    int     조회구분 : 전체(0), 매도(1), 매수(2)
    sygb = 0    '    int     신용구분 : 전체(0), 신용(1), 담보(2)
    dlgb = 2    '    int     미체구분 : 전체(0), 체결(1), 미체결(2)
    sort = 2    '    int     정렬구분 : 역순(1), 정순(2)
    nkey = ""   '    char    18  [Optional] 연속조회키

    If AxIBK.TR1211(TRKey.TR1211, acno, pswd, code, dsgb, sygb, dlgb, sort, nkey) = False Then
        MsgBox AxIBK.GetLastErrMsg
    End If

End Sub

Public Sub btnIBKSTR1211_Click()
    '주식:주식 체결내역 조회-전체(TR1211)
    Dim Key, acno, pswd, code, dsgb, sygb, dlgb, sort, nkey

    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)
    
    HONAQ003Count = -1
    HONAQ003Count_T = 0
    isSelectHONAQ003 = False
    
    acno = Sheets("모듈").Cells(6, 12).Value    '    char    11  계좌번호
    pswd = Sheets("모듈").Cells(7, 12).Value    '    char    8   계좌비밀번호
    code = ""   '    char    7   [Option]조회종목번호
    dsgb = 0    '    int     조회구분 : 전체(0), 매도(1), 매수(2)
    sygb = 0    '    int     신용구분 : 전체(0), 신용(1), 담보(2)
    dlgb = 0    '    int     미체구분 : 전체(0), 체결(1), 미체결(2)
    sort = 2    '    int     정렬구분 : 역순(1), 정순(2)
    nkey = ""   '    char    18  [Optional] 연속조회키

    If AxIBK.TR1211(TRKey.TR1211, acno, pswd, code, dsgb, sygb, dlgb, sort, nkey) = False Then
        MsgBox AxIBK.GetLastErrMsg
    End If

End Sub

Public Sub btnIBKSTR1221_Click()
    '주식:잔고조회(TR1221)
    Dim Key, acno, pswd, allf, nkey

    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)
    
    HONAQ001Count = -1
    HONAQ001Count_T = 0
    
    acno = Sheets("모듈").Cells(6, 12).Value    '    char    11  계좌번호
    pswd = Sheets("모듈").Cells(7, 12).Value    '    char    8   계좌비밀번호
    allf = 1    '    int     신용구분 : 신용제외(0), 모두(1), 신용(2)
    nkey = ""   '    char    18  [Optional] 연속조회키

    'MsgBox AxIBK.GetAccounts()
    'MsgBox AxIBK.GetUserID()
    
    'MsgBox AxIBK.GetUserName()


    If AxIBK.TR1221(TRKey.TR1221, acno, pswd, allf, nkey) = False Then
        MsgBox AxIBK.GetLastErrMsg
    End If

End Sub

Public Sub btnLogin_Click()
  
    UserID = Sheets("모듈").Cells(1, 2).Value
    UserPass = Sheets("모듈").Cells(2, 2).Value
    ReplayGubun = Sheets("모듈").Cells(4, 2).Value
    OrderGubun = Sheets("모듈").Cells(4, 5).Value
    
    QuoteStockHStartPos = CInt(Sheets("모듈").Cells(1, 7).Value)
    QuoteStockVStartPos = CInt(Sheets("모듈").Cells(2, 7).Value)
    QuoteStockItemCount = CInt(Sheets("모듈").Cells(3, 7).Value)

    QuoteFutureHStartPos = CInt(Sheets("모듈").Cells(4, 7).Value)
    QuoteFutureVStartPos = CInt(Sheets("모듈").Cells(5, 7).Value)
    QuoteFutureItemCount = CInt(Sheets("모듈").Cells(6, 7).Value)

    LoginProcess

End Sub

Public Sub btnLoginSK_Click()
    ' 로그인

'    Dim nResult As Integer
'    nResult = WapiConnectMngCtrl.LoginServer()
'
'
'    If nResult = 0 Then
'        Sheets("모듈").Cells(8, 9).Value = "0(정상)"
'
'        ' 계좌번호 조회
'
'    ElseIf nResult = 1 Then
'        Sheets("모듈").Cells(8, 9).Value = "1(실패, UserID가 존재 하지 않음)"
'    ElseIf nResult = 2 Then
'        Sheets("모듈").Cells(8, 9).Value = "2(실패, Password가 틀렸음)"
'    ElseIf nResult = 3 Then
'        Sheets("모듈").Cells(8, 9).Value = "3(실패, Password 5회 실수 로긴 거부)"
'    ElseIf nResult = 4 Then
'        Sheets("모듈").Cells(8, 9).Value = "4(실패, 계좌 없는 사용자(비등록자))"
'    ElseIf nResult = 5 Then
'        Sheets("모듈").Cells(8, 9).Value = "5(실패, 유효하지 않은 사용자)"
'    ElseIf nResult = 6 Then
'        Sheets("모듈").Cells(8, 9).Value = "6(실패, System(DB) Error)"
'    ElseIf nResult = 7 Then
'        Sheets("모듈").Cells(8, 9).Value = "7(실패, 공인인증 실패)"
'    ElseIf nResult = 100 Then
'        Sheets("모듈").Cells(8, 9).Value = "100(실패, 해당 XML파일이 없음)"
'    Else
'        Sheets("모듈").Cells(8, 9).Value = "실패 "
'    End If
End Sub

Public Sub btnSKAccountList_Click()
'    '계좌리스트 조회(09004)
'
'    Dim strTrNo As String
'    Dim strCodeNo As String
'    Dim nPoint As Long
'
'    strTrNo = "09004"
'    strCodeNo = Sheets("모듈").Cells(1, 9).Value
'
'    nPoint = WapiCommunityMngCtrl.InitTR(strTrNo)
'
'    If nPoint > 0 Then
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "09004.사용자아이디", strCodeNo, 0)
'        Call WapiCommunityMngCtrl.SendTr(strTrNo, 0)
'    End If
    
End Sub

Public Sub btnSKRegistReal_Click()
'' 리얼등록
'
'    Dim strRealPacket As String
'    Dim strPrimaryKey As String
'
'    strRealPacket = "XX1"      ' 트랜정보
'    strPrimaryKey = Sheets("모듈").Cells(1, 9).Value      ' 입력정보
'
'    Call WapiCommunityMngCtrl.RegistReal(strRealPacket, strPrimaryKey)

End Sub

Public Sub btnSKReleaseReal_Click()
'' 리얼해지
'
'    Dim strRealPacket As String
'    Dim strPrimaryKey As String
'
'    strRealPacket = "XX1"      ' 리얼패킷
'    strPrimaryKey = Sheets("모듈").Cells(1, 9).Value      ' 입력정보
'
'    Call WapiCommunityMngCtrl.ReleaseReal(strRealPacket, strPrimaryKey)
End Sub

Public Sub btnSKStock96007_Click()
'    '주식:잔고조회(96007)
'
'    Dim nPoint As Long
'
'    HONAQ001Count = -1
'
'    nPoint = WapiCommunityMngCtrl.InitTR("96007")
'
'    If nPoint > 0 Then
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96007.계좌번호", Sheets("모듈").Cells(4, 9).Value, 0)
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96007.비밀번호", Sheets("모듈").Cells(5, 9).Value, 0)
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96007.주문대리인구분", "0", 0)   '0:일반, 1:주문대리인
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96007.대표계좌번호", "", 0) '주문대리인구분 1로 선택했을 경우 주문대리인 계좌번호 입력
'        Call WapiCommunityMngCtrl.SendTr("96007", 0)
'    End If
End Sub

Public Sub btnSKStock96012_Click()
'    '주식:주문체결/전체 조회(96012)
'
'    Dim nPoint As Long
'
'    HONAQ003Count = -1
'    isSelectHONAQ003 = False
'
'    nPoint = WapiCommunityMngCtrl.InitTR("96012")
'
'    If nPoint > 0 Then
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96012.계좌번호", Sheets("모듈").Cells(4, 9).Value, 0)
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96012.비밀번호", Sheets("모듈").Cells(5, 9).Value, 0)
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96012.종목코드", "", 0)
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96012.매매구분", "0", 0)   '0:전체, 1:매도, 2:매수
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96012.주문대리인구분", "0", 0)   '0:일반, 1:주문대리인
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96012.대표계좌번호", "", 0) '주문대리인구분 1로 선택했을 경우 주문대리인 계좌번호 입력
'        Call WapiCommunityMngCtrl.SendTr("96012", 0)
'    End If
    
End Sub

Public Sub btnSKStock96013_Click()
'    '주식:미체결 조회(96013)
'
'    Dim nPoint As Long
'
'    HONAQ003Count = -1
'    isSelectHONAQ003 = False
'
'    nPoint = WapiCommunityMngCtrl.InitTR("96013")
'
'    If nPoint > 0 Then
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96013.계좌번호", Sheets("모듈").Cells(4, 9).Value, 0)
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96013.비밀번호", Sheets("모듈").Cells(5, 9).Value, 0)
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96013.종목코드", "", 0)
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96013.매매구분", "0", 0)   '0:전체, 1:매도, 2:매수
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96013.주문대리인구분", "0", 0)   '0:일반, 1:주문대리인
'        Call WapiCommunityMngCtrl.SetTrData(nPoint, "96013.대표계좌번호", "", 0) '주문대리인구분 1로 선택했을 경우 주문대리인 계좌번호 입력
'        Call WapiCommunityMngCtrl.SendTr("96013", 0)
'    End If

End Sub

Public Sub btnSKStockBuy_Click()
'    '주식매수주문
'    Dim sIssueCode As String    '코드
'    Dim sAccount As String    '계좌번호
'    Dim sPassword As String    '비밀번호
'    Dim sOrderType As String    '주문조건
'    Dim sCallType As String    '호가유형
'    Dim sQty As String    '매매수량
'    Dim sPrice As String    '매매가격
'    Dim i As Integer
'    Dim nPoint As Long
'
'    AccountPass = Sheets("모듈").Cells(5, 9).Value
'
'    sAccount = Sheets("모듈").Cells(4, 9).Value
'    sPassword = AccountPass
'
'    sCallType = "00" '00: 지정가, 03: 시장가, 05: 조건부지정가, 31: 최유리지정가, 32: 최우선지정가, 61: 장전시간외종가, 99: 장후시간외종가, 81: 시간외단일가매매
'
'    nPoint = WapiCommunityMngCtrl.InitTR("96062")
'
'    '주식매수주문
'    If nPoint > 0 Then
'        For i = 1 To QuoteStockItemCount
'            sIssueCode = Right(Sheets("가격수령").Cells(3, i + QuoteStockHStartPos - 1).Value, 6)
'            sPrice = Sheets("가격수령").Cells(29, i + QuoteStockHStartPos - 1).Value
'            sQty = Sheets("가격수령").Cells(30, i + QuoteStockHStartPos - 1).Value
'            'sPrice = "0" '현재가로 테스트함
'            'sQty = 1
'
'            If (sPrice <> "0") And (sQty <> "0") Then
'            'If (sQty <> "0") Then
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.ord_tp", "2", 0)    '1:매도, 2:매수, 3:정정, 4:취소
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.계좌번호", sAccount, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.비밀번호", AccountPass, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.종목구분", "1", 0)  '1:장내, 2:장외, 3:제3시장
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.종목코드", sIssueCode, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.매매구분", sCallType, 0) '00: 지정가, 03: 시장가, 05: 조건부지정가, 31: 최유리지정가, 32: 최우선지정가, 61: 장전시간외종가, 99: 장후시간외종가, 81: 시간외단일가매매
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.체결구분", "0", 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문수량", sQty, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문단가", sPrice, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.원주문번호", "", 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문대리인구분", "0", 0)    '0또는 space:일반, 1:주문대리인
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.대표계좌번호", "", 0)   '주문대리인구분 1로 선택했을 경우 주문대리인 계좌번호 입력, 0일때는 space
'
'                Call WapiCommunityMngCtrl.SendTr("96062", 0)
'                WaitFor (0.1)
'            End If
'
'        Next
'    End If
    
End Sub

Public Sub btnSKStockCancel_Click()
'    '주식매도주문
'    Dim sIssueCode As String    '코드
'    Dim sAccount As String    '계좌번호
'    Dim sPassword As String    '비밀번호
'    Dim sOrgOrdNum As String    '원주문번호
'    Dim sQty As String    '매수수량
'    Dim iRow As Integer
'    Dim iCol As Integer
'    Dim nPoint As Long
'
'    AccountPass = Sheets("모듈").Cells(5, 9).Value
'
'    sAccount = Sheets("모듈").Cells(4, 9).Value
'    sPassword = AccountPass
'
'
'    nPoint = WapiCommunityMngCtrl.InitTR("96062")
'
'    If nPoint > 0 Then
'        For iRow = 0 To UBound(HONAQ003Data, 1) - 1
'            sIssueCode = HONAQ003Data(iRow, 10)
'            sOrgOrdNum = CStr(CInt(HONAQ003Data(iRow, 7)))
'            sQty = CStr(CInt(HONAQ003Data(iRow, 4)))
'
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.ord_tp", "4", 0)    '1:매도, 2:매수, 3:정정, 4:취소
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.계좌번호", sAccount, 0)
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.비밀번호", sPassword, 0)
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.종목구분", "1", 0)  '1:장내, 2:장외, 3:제3시장
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.종목코드", sIssueCode, 0)
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.매매구분", "00", 0) '00: 지정가, 03: 시장가, 05: 조건부지정가, 31: 최유리지정가, 32: 최우선지정가, 61: 장전시간외종가, 99: 장후시간외종가, 81: 시간외단일가매매
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.체결구분", "0", 0)
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문수량", sQty, 0)
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문단가", "0", 0)
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.원주문번호", sOrgOrdNum, 0)
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문대리인구분", "0", 0)    '0또는 space:일반, 1:주문대리인
'            Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.대표계좌번호", "", 0)   '주문대리인구분 1로 선택했을 경우 주문대리인 계좌번호 입력, 0일때는 space
'
'            Call WapiCommunityMngCtrl.SendTr("96062", 0)
'            WaitFor (0.1)
'        Next
'    End If
    
End Sub

Public Sub btnSKStockSell_Click()
'    '주식매도주문
'    Dim sIssueCode As String    '코드
'    Dim sAccount As String    '계좌번호
'    Dim sPassword As String    '비밀번호
'    Dim sOrderType As String    '주문조건
'    Dim sCallType As String    '호가유형
'    Dim sQty As String    '매매수량
'    Dim sPrice As String    '매매가격
'    Dim i As Integer
'    Dim nPoint As Long
'
'    AccountPass = Sheets("모듈").Cells(5, 9).Value
'
'    sAccount = Sheets("모듈").Cells(4, 9).Value
'    sPassword = AccountPass
'
'    sCallType = "00" '00: 지정가, 03: 시장가, 05: 조건부지정가, 31: 최유리지정가, 32: 최우선지정가, 61: 장전시간외종가, 99: 장후시간외종가, 81: 시간외단일가매매
'
'    '주식매도주문
'    nPoint = WapiCommunityMngCtrl.InitTR("96062")
'
'    If nPoint > 0 Then
'        For i = 1 To QuoteStockItemCount
'            sIssueCode = Right(Sheets("가격수령").Cells(3, i + QuoteStockHStartPos - 1).Value, 6)
'            sPrice = Sheets("가격수령").Cells(31, i + QuoteStockHStartPos - 1).Value
'            sQty = Sheets("가격수령").Cells(32, i + QuoteStockHStartPos - 1).Value
'            'sPrice = "0" '현재가로 테스트함
'            'sQty = 1
'
'            If (sPrice <> "0") And (sQty <> "0") Then
'            'If (sQty <> "0") Then
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.ord_tp", "1", 0)    '1:매도, 2:매수, 3:정정, 4:취소
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.계좌번호", sAccount, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.비밀번호", AccountPass, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.종목구분", "1", 0)  '1:장내, 2:장외, 3:제3시장
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.종목코드", sIssueCode, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.매매구분", sCallType, 0) '00: 지정가, 03: 시장가, 05: 조건부지정가, 31: 최유리지정가, 32: 최우선지정가, 61: 장전시간외종가, 99: 장후시간외종가, 81: 시간외단일가매매
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.체결구분", "0", 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문수량", sQty, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문단가", sPrice, 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.원주문번호", "", 0)
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.주문대리인구분", "0", 0)    '0또는 space:일반, 1:주문대리인
'                Call WapiCommunityMngCtrl.SetTrData(nPoint, "96062.대표계좌번호", "", 0)   '주문대리인구분 1로 선택했을 경우 주문대리인 계좌번호 입력, 0일때는 space
'
'                Call WapiCommunityMngCtrl.SendTr("96062", 0)
'                WaitFor (0.1)
'            End If
'
'        Next
'    End If
    
End Sub

Public Sub btnSKVersionExec_Click()
'    '버전파일 실행
'    Dim sProgPath As String
'    Dim nResult As Double
'
'    sProgPath = Trim(Sheets("모듈").Cells(7, 9).Value)
'    'sProgPath = "cmd.exe /c C:\SKS\Wapi_T\CommVerMng.exe"
'
'    nResult = Shell(sProgPath, vbNormalFocus)
'
'    'MsgBox CStr(nResult)
End Sub

Public Sub btnStockHONAQ001_Click()
    Dim sIssueCode As String     '계좌번호
    Dim sPassword As String     '비밀번호
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value
    
    sIssueCode = AccountList(0, 0)
    sPassword = AccountPass
    
    CreateBizObject
    
    HONAQ001Count = -1
    
    Sheets("모듈").OrderObject.SendReceiveEx sIssueCode, sPassword, "HONAQ001"   '주식잔고조회

End Sub

Public Sub btnStockHONAQ003_Click()
    Dim sIssueCode As String     '계좌번호
    Dim sPassword As String     '비밀번호
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value
    
    sIssueCode = AccountList(0, 0)
    sPassword = AccountPass
    
    CreateBizObject
    
    HONAQ003Count = -1
    isSelectHONAQ003 = False
    
    Sheets("모듈").OrderObject.SendReceiveEx sIssueCode, sPassword, "HONAQ003", "2"   '체결내역조회 조회구분 0:전체(체결,미체결), 1:체결, 2:미체결
End Sub

Private Sub CommandButton1_Click()
    Dim sItemCodes As String
    Dim iCol As Integer
    
    For iCol = QuoteFutureHStartPos To QuoteFutureItemCount + QuoteFutureHStartPos - 1
        sItemCodes = sItemCodes & Sheets("가격수령").Cells(3, iCol).Value & ","
        If iCol = 211 Then
            sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
        End If
    Next
    
    MsgBox sItemCodes

End Sub

Private Sub CommandButton2_Click()
    Dim sEndTime As String
    
    sEndTime = Sheets("모듈").Cells(1, 5).Value
    IsCancelPress = True
    
    Application.OnTime EarliestTime:=TimeValue(sEndTime), Procedure:="Finished", Schedule:=False
End Sub

Public Sub FutureBatchBuy_Click()
    Dim sIssueCode As String    '코드
    Dim sAccount As String    '계좌번호
    Dim sPassword As String    '비밀번호
    Dim sOrderType As String    '주문조건
    Dim sCallType As String    '호가유형
    Dim sQty As String    '매수수량
    Dim sPrice As String    '매수가격
    Dim i As Integer

    AccountPass = Sheets("모듈").Cells(3, 2).Value
    
    sAccount = AccountListFO(0, 0)
    sPassword = AccountPass

    sOrderType = "1" '"1" 일반, "2" IOC, "3" FOK
    sCallType = "00" '"00" 지정가, "03" 시장가, "05" 조건부, "06" 최유리, "07" 최우선

    
    'CreateBizObject
    
    For i = 1 To QuoteFutureItemCount
        sIssueCode = Sheets("가격수령").Cells(3, i + QuoteFutureHStartPos - 1).Value
        sPrice = Sheets("가격수령").Cells(29, i + QuoteFutureHStartPos - 1).Value
        sQty = Sheets("가격수령").Cells(30, i + QuoteFutureHStartPos - 1).Value
        If (sPrice <> "0") And (sQty <> "0") Then
            With Sheets("모듈")
                .OrderObject.SendBuyFO sAccount, sPassword, sIssueCode, sOrderType, sCallType, sQty, sPrice
                WaitFor (0.01)
            End With
        End If
    Next
End Sub

Public Sub FutureBatchCancel_Click()
    Dim sIssueCode As String    '코드
    Dim sAccount As String    '계좌번호
    Dim sPassword As String    '비밀번호
    Dim sOrgOrdNum As String    '원주문번호
    Dim sQty As String    '매수수량
    Dim iRow As Integer
    Dim iCol As Integer
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value
    
    sAccount = AccountListFO(0, 0)
    sPassword = AccountPass

    'CreateBizObject
    
    
    For iRow = 0 To UBound(HONBQ003Data, 1) - 1
        sIssueCode = HONBQ003Data(iRow, 4)
        sOrgOrdNum = HONBQ003Data(iRow, 1)
        sQty = HONBQ003Data(iRow, 21)
        With Sheets("모듈")
            .OrderObject.SendCancelFO sAccount, sPassword, sIssueCode, sOrgOrdNum, sQty
            WaitFor (0.01)
        End With
    Next
End Sub

Public Sub FutureBatchSell_Click()
    Dim sIssueCode As String    '코드
    Dim sAccount As String    '계좌번호
    Dim sPassword As String    '비밀번호
    Dim sOrderType As String    '주문조건
    Dim sCallType As String    '호가유형
    Dim sQty As String    '매수수량
    Dim sPrice As String    '매수가격
    Dim i As Integer
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value

    sAccount = AccountListFO(0, 0)
    sPassword = AccountPass

    sOrderType = "1" '"1" 일반, "2" IOC, "3" FOK
    sCallType = "00" '"00" 지정가, "03" 시장가, "05" 조건부, "06" 최유리, "07" 최우선

    
    'CreateBizObject
    
    For i = 1 To QuoteFutureItemCount
        sIssueCode = Sheets("가격수령").Cells(3, i + QuoteFutureHStartPos - 1).Value
        sPrice = Sheets("가격수령").Cells(31, i + QuoteFutureHStartPos - 1).Value
        sQty = Sheets("가격수령").Cells(32, i + QuoteFutureHStartPos - 1).Value
        If (sPrice <> "0") And (sQty <> "0") Then
            With Sheets("모듈")
                .OrderObject.SendSellFO sAccount, sPassword, sIssueCode, sOrderType, sCallType, sQty, sPrice
                WaitFor (0.01)
            End With
        End If
    Next
End Sub

Private Sub IBKSConnector1_OnLogin(ByVal bLogin As Boolean, ByVal msg As String)

End Sub

Private Sub OrderBatch_Click()
    Dim iCount As Integer
    
    Sheets("모듈").QuotePause_Click
    
    WaitFor (1)
    
    'Application.OnTime Now + TimeValue("00:00:10"), "매매계속"
    
'    WaitFor (1)
    
    Sheets("모듈").StockBatchBuy_Click
'    WaitFor (1)
'    Sheets("모듈").FutureBatchBuy_Click
    WaitFor (1)
    Sheets("모듈").StockBatchSell_Click
'    WaitFor (1)
'    Sheets("모듈").FutureBatchSell_Click
    
    '주문 후 시간 대기 한다.
    WaitFor (3)
    
    '조회 후 데이터를 받을때까지 대기 한다.
    '미체결내역조회
    
    Do While True
    '   주식취소
        Sheets("모듈").btnStockHONAQ003_Click
        iCount = 0
        Do While iCount < 10
            WaitFor (1)
            If (HONAQ003Count >= 0) Then
                Exit Do
            End If
            iCount = iCount + 1
        Loop
        
        WaitFor (1)
        
    '   취소전송
        If (HONAQ003Count > 0) Then
            Sheets("모듈").StockBatchCancel_Click
        End If
        
        ' 선물을 운용 할 때 HONBQ003Count도 따로 체크 해야함.
        If (HONAQ003Count <= 0) Then
            Exit Do
        End If
        
    '   선물 취소
    '    Sheets("모듈").btnFOHONBQ003_Click
    '    iCount = 0
    '    Do While iCount < 10
    '        WaitFor (1)
    '        If HONBQ003Count >= 0 Then
    '            Exit Do
    '        End If
    '        iCount = iCount + 1
    '    Loop
            
    '    WaitFor (1)
    '
    '   취소전송
    '    If (HONBQ003Count > 0) Then
    '        Sheets("모듈").FutureBatchCancel_Click
    '    End If
    
    Loop
    
'   취소후 3초 대기
    WaitFor (3)
    
    '잔고조회
    Sheets("모듈").btnStockHONAQ001_Click
    iCount = 0
    Do While iCount < 10
        WaitFor (1)
        If HONAQ001Count >= 0 Then
            Exit Do
        End If
        iCount = iCount + 1
    Loop

    WaitFor (1)
    
'    Sheets("모듈").btnFOHONBQ001_Click
'    iCount = 0
'    Do While iCount < 10
'        WaitFor (1)
'        If HONBQ001Count >= 0 Then
'            Exit Do
'        End If
'        iCount = iCount + 1
'    Loop
    
    
    매매계속
    
    WaitFor (1)
    
    '시세재시작
    Sheets("모듈").QuoteStart_Click

    HONAQ001Count = -1
    HONBQ001Count = -1
    HONAQ003Count = -1
    HONBQ003Count = -1
End Sub

Private Sub OrderFirst_Click()
    Dim iCol As Integer
    
    Sheets("모듈").QuotePause_Click
    
'    WaitFor (1)
    
    'Application.OnTime Now + TimeValue("00:00:10"), "매매계속"
    
    WaitFor (1)
    
    '잔고엑셀초기화
    For iCol = QuoteStockHStartPos To QuoteStockItemCount + QuoteStockHStartPos - 1
        Sheets("가격수령").Cells(15, iCol).Value = 0
        Sheets("가격수령").Cells(16, iCol).Value = 0
    Next
    
    
    첫매매
    
    WaitFor (1)

    '시세재시작
    Sheets("모듈").QuoteStart_Click

    HONAQ001Count = -1
    HONBQ001Count = -1
    HONAQ003Count = -1
    HONBQ003Count = -1
End Sub

Private Sub OrderObject_BizServerClose()
    QuotePause_Click
    WaitFor (3)
    QuoteStop_Click
    WaitFor (3)
    ThisWorkbook.Save
    Application.Quit
End Sub

Private Sub OrderObject_BizServerInit(ByVal bsData As String)
    'MsgBox bsData
    Sheets("모듈").OrderObject.SetMainDir "C:\ExcelAPI"
    AccountSeq = 0

    Sheets("모듈").OrderObject.OpenComplete
End Sub

Private Sub OrderObject_ReceiveAccountList(ByVal bsAccount As String, ByVal bsAccountName As String)
On Error GoTo ErrMsg
    DoEvents
    
    If StrComp(bsAccount, "----------") <> 0 Then
        ReDim Preserve AccountList(1, AccountSeq)
        AccountList(0, AccountSeq) = bsAccount
        AccountList(1, AccountSeq) = bsAccountName
        AccountSeq = AccountSeq + 1
    End If
    
    If StrComp(bsAccount, "----------") = 0 Then
        IsLogined = True
        AccountSeq = 0
        
        Sheets("모듈").OrderObject.GetItemList "kospi", "stocks", True
        DoEvents
        Sheets("모듈").OrderObject.GetItemList "kosdaq", "stocks", False
        
    End If

    Exit Sub
ErrMsg:
    If (Err.Number = 50290) Then
        MsgBox "ReceiveAccountList 오류번호 : " & Err.Number & vbCr & _
       "오류내용 : " & Err.Description, vbCritical, "오류"
    End If
    If (Err.Number <> 0) And (Err.Number <> 50290) Then
        MsgBox "ReceiveAccountList 오류번호 : " & Err.Number & vbCr & _
       "오류내용 : " & Err.Description, vbCritical, "오류"
    End If
    Exit Sub
End Sub

Private Sub OrderObject_ReceiveAccountListFO(ByVal bsAccount As String, ByVal bsAccountName As String)
On Error GoTo ErrMsg
    DoEvents
    
    If StrComp(bsAccount, "----------") <> 0 Then
        ReDim Preserve AccountListFO(1, AccountSeqFO)
        AccountListFO(0, AccountSeqFO) = bsAccount
        AccountListFO(1, AccountSeqFO) = bsAccountName
        AccountSeqFO = AccountSeqFO + 1
    End If

    If StrComp(bsAccount, "----------") = 0 Then
        IsLogined = True
        AccountSeqFO = 0
        
        DoEvents
        Sheets("모듈").OrderObject.GetItemList "kospi200", "futures", True
        DoEvents
        Sheets("모듈").OrderObject.GetItemList "kospi200", "options", False
        DoEvents
        Sheets("모듈").OrderObject.GetItemList "mini_kospi200", "futures", True
        DoEvents
        Sheets("모듈").OrderObject.GetItemList "mini_kospi200", "options", False
        DoEvents
        Sheets("모듈").OrderObject.GetItemList "stock", "futures", True
        'DoEvents
        'Sheets("모듈").OrderObject.GetItemList "stock", "options", False
    
    End If

    Exit Sub
ErrMsg:
    If (Err.Number = 50290) Then
        MsgBox "ReceiveAccountListFO 오류번호 : " & Err.Number & vbCr & _
       "오류내용 : " & Err.Description, vbCritical, "오류"
    End If
    If (Err.Number <> 0) And (Err.Number <> 50290) Then
        MsgBox "ReceiveAccountListFO 오류번호 : " & Err.Number & vbCr & _
       "오류내용 : " & Err.Description, vbCritical, "오류"
    End If
    Exit Sub

End Sub

Private Sub OrderObject_ReceiveExecute(ByVal nType As Long, ByVal bsData As String)
   
    Select Case nType
        Case 1
            'Sheets("모듈").Cells(35, 1).Value = "Config Data: " & bsData
            
            ParseJSONToSet bsData
        Case 2
            'Sheets("모듈").Cells(36, 1).Value = "Login Process: " & bsData
            
            ' bsData 값에 따라 다른 로직 실행
            Select Case bsData
                Case "1"
                    Sheets("모듈").btnLogin_Click
                    ' 여기에 로그인 성공시 실행할 추가 코드를 넣습니다.
                Case "2"
                    Sheets("모듈").btnIBKSInit_Click
                    WaitFor (1)
                    Sheets("모듈").btnIBKSConnect_Click
                    WaitFor (1)
                    Sheets("모듈").btnLogin_Click
                    ' 여기에 로그인 실패시 실행할 추가 코드를 넣습니다.
                Case "3"
                    ' 여기에 로그인 재시도시 실행할 추가 코드를 넣습니다.
                Case Else
                    ' 여기에 알 수 없는 bsData 값에 대한 처리 코드를 넣습니다.
            End Select
            SendDataToMain GetHandleValue, MAIN_SECURITY_READY, ""
        Case 3
            'Sheets("모듈").Cells(37, 1).Value = "Account Process"
            Sheets("모듈").btnAccount_Click
        Case 4
            'Sheets("모듈").Cells(38, 1).Value = "Quote STK Process"
            Sheets("모듈").QuoteCallStock_Click
        Case 5
            'Sheets("모듈").Cells(39, 1).Value = "Quote FO Process"
            Sheets("모듈").QuoteCallFutures_Click
        Case 11
            'Sheets("모듈").Cells(40, 1).Value = "Macro Start Process"
            첫매매
        Case 12
            'Sheets("모듈").Cells(41, 1).Value = "Macro Next Process"
            매매계속
        Case 13
            'Sheets("모듈").Cells(42, 1).Value = "Macro Continue Process"
            연속매매
        Case 14
            'Sheets("모듈").Cells(43, 1).Value = "Trade Data"
            If OrderGubun = "HINT" Then
                StockHONAQ003Search
            ElseIf OrderGubun = "SK" Then
                Sheets("모듈").btnSKStock96013_Click
            ElseIf OrderGubun = "IBKS" Then
                Sheets("모듈").btnIBKSTR1211_Click
            End If
        Case 15
            'Sheets("모듈").Cells(44, 1).Value = "ES Data"
            If OrderGubun = "HINT" Then
                Sheets("모듈").btnStockHONAQ001_Click
            ElseIf OrderGubun = "SK" Then
                Sheets("모듈").btnSKStock96007_Click
            ElseIf OrderGubun = "IBKS" Then
                Sheets("모듈").btnIBKSTR1221_Click
            End If
        Case 16
            CommandButton2_Click
        Case Else
            'Sheets("모듈").Cells(45, 1).Value = "Nothing To do"
    End Select
    
End Sub


Private Sub OrderObject_ReceiveItemList(ByVal bsMarketType As String, ByVal bsItemType As String, ByVal nRowSize As Long, ByVal nColSize As Long)
    '종목리스트 처리
'    Dim Data() As String
'    Dim iRow As Integer
'    Dim iCol As Integer
'    Dim strData As String
'
'    Application.ScreenUpdating = False
'    'Application.DisplayStatusBar = False
'    Application.Calculation = xlCalculationManual
'    'Application.EnableEvents = False
'    ActiveSheet.DisplayPageBreaks = False
'
'
'    If (nColSize < 0) Or (nRowSize < 0) Then
'        Exit Sub
'    End If
'
'    ReDim Data(nRowSize, nColSize) As String
'
'    Sheets("모듈").OrderObject.GetRepeatIResult Data
'
'    If (bsMarketType = "kospi") And (bsItemType = "stocks") Then
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'                If (iCol = 0) Then '종목코드
'                    Sheets("주식종목").Cells(iRow + 3, 1).Value = strData
'                End If
'                If (iCol = 1) Then '종목명
'                    Sheets("주식종목").Cells(iRow + 3, 2).Value = strData
'                End If
'            Next
'        Next
'    End If
'
'    If (bsMarketType = "kosdaq") And (bsItemType = "stocks") Then
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'                If (iCol = 0) Then '종목코드
'                    Sheets("주식종목").Cells(iRow + 3, 4).Value = strData
'                End If
'                If (iCol = 1) Then '종목명
'                    Sheets("주식종목").Cells(iRow + 3, 5).Value = strData
'                End If
'            Next
'        Next
'    End If
'
'    If (bsMarketType = "kospi200") And (bsItemType = "futures") Then
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'                If (iCol = 0) Then '종목코드
'                    Sheets("선옵종목").Cells(iRow + 3, 1).Value = strData
'                End If
'                If (iCol = 1) Then '종목명
'                    Sheets("선옵종목").Cells(iRow + 3, 2).Value = strData
'                End If
'            Next
'        Next
'    End If
'
'    If (bsMarketType = "kospi200") And (bsItemType = "options") Then
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'                If (iCol = 0) Then '종목코드
'                    Sheets("선옵종목").Cells(iRow + 3, 4).Value = strData
'                End If
'                If (iCol = 1) Then '종목명
'                    Sheets("선옵종목").Cells(iRow + 3, 5).Value = strData
'                End If
'            Next
'        Next
'    End If
'
'    If (bsMarketType = "mini_kospi200") And (bsItemType = "futures") Then
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'                If (iCol = 0) Then '종목코드
'                    Sheets("선옵종목").Cells(iRow + 3, 7).Value = strData
'                End If
'                If (iCol = 1) Then '종목명
'                    Sheets("선옵종목").Cells(iRow + 3, 8).Value = strData
'                End If
'            Next
'        Next
'    End If
'
'    If (bsMarketType = "mini_kospi200") And (bsItemType = "options") Then
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'                If (iCol = 0) Then '종목코드
'                    Sheets("선옵종목").Cells(iRow + 3, 10).Value = strData
'                End If
'                If (iCol = 1) Then '종목명
'                    Sheets("선옵종목").Cells(iRow + 3, 11).Value = strData
'                End If
'            Next
'        Next
'    End If
'
'    If (bsMarketType = "stock") And (bsItemType = "futures") Then
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'                If (iCol = 0) Then '종목코드
'                    Sheets("선옵종목").Cells(iRow + 3, 13).Value = strData
'                End If
'                If (iCol = 1) Then '종목명
'                    Sheets("선옵종목").Cells(iRow + 3, 14).Value = strData
'                End If
'            Next
'        Next
'    End If
'
'    If (bsMarketType = "stock") And (bsItemType = "options") Then
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'                If (iCol = 0) Then '종목코드
'                    Sheets("선옵종목").Cells(iRow + 3, 16).Value = strData
'                End If
'                If (iCol = 1) Then '종목명
'                    Sheets("선옵종목").Cells(iRow + 3, 17).Value = strData
'                End If
'            Next
'        Next
'    End If
'
'    Application.ScreenUpdating = screenUpdateState
'    'Application.DisplayStatusBar = statusBarState
'    'Application.Calculation = calcState
'    'Application.EnableEvents = eventsState
'    'ActiveSheet.DisplayPageBreaks = displayPageBreakState

End Sub

Private Sub OrderObject_ReceiveMessage(ByVal bsData As String)
    '모든 처리응답메세지 처리
    'Sheets("모듈").Cells(8, 2).Value = bsData

End Sub

Private Sub OrderObject_ReceiveNotice(ByVal nRowSize As Long, ByVal nColSize As Long)
    '주문 노티 수신
'    Dim Data() As String
'    If (nColSize < 0) Or (nRowSize < 0) Then
'        Exit Sub
'    End If
'
'    ReDim Data(nRowSize, nColSize) As String
'
'    Sheets("모듈").OrderObject.GetRepeatNResult Data
'
'    Dim iRow As Integer
'    Dim iCol As Integer
'
'    Dim strData As String
'
'    For iRow = 0 To UBound(Data, 1) - 1
        
'        If (Data(iRow, 0) = "매수체결") Or (Data(iRow, 0) = "매도체결") Then
'            Sheets("모듈").Range("C36", "I36").Insert Shift:=xlDown
'
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'
'                Select Case iCol
'                Case 0 '구분
'                    Sheets("모듈").Cells(36, 3).Value = strData
'                Case 1 '번호
'                    Sheets("모듈").Cells(36, 4).Value = strData
'                Case 2 '계좌번호
'                    Sheets("모듈").Cells(36, 5).Value = strData
'                Case 3 '종목명
'                    Sheets("모듈").Cells(36, 7).Value = strData
'                Case 4 '체결가
'                    Sheets("모듈").Cells(36, 8).Value = strData
'                Case 5 '체결수량
'                    Sheets("모듈").Cells(36, 9).Value = strData
'                Case 6 '종목코드
'                    Sheets("모듈").Cells(36, 6).Value = strData
'                End Select
'            Next
'        End If
        
'        For iCol = 0 To UBound(Data, 2)
'            Sheets("모듈").Range("C36", "I36").Insert Shift:=xlDown
'            strData = Data(iRow, iCol)
'            Sheets("모듈").Cells(36, iCol + 3).Value = strData
            
'            If (Data(iRow, 0) = "매수체결") And (Data(iRow, 0) = "매도체결") Then
'
'
'                Sheets("모듈").Range("C36", "I36").Insert Shift:=xlDown
'
'                strData = Data(iRow, iCol)
'
'                Sheets("모듈").Cells(36, 3).Value = strData
    
'                Select Case iCol
'                Case 0 '구분
'                    Sheets("모듈").Cells(36, 3).Value = strData
'                Case 1 '번호
'                    Sheets("모듈").Cells(36, 4).Value = strData
'                Case 2 '계좌번호
'                    Sheets("모듈").Cells(36, 5).Value = strData
'                Case 3 '종목명
'                    Sheets("모듈").Cells(36, 7).Value = strData
'                Case 4 '체결가
'                    Sheets("모듈").Cells(36, 8).Value = strData
'                Case 5 '체결수량
'                    Sheets("모듈").Cells(36, 9).Value = strData
'                Case 6 '종목코드
'                    Sheets("모듈").Cells(36, 6).Value = strData
'                End Select
                
'                Sheets("가격수령").Range("C36", "I36").Insert Shift:=xlDown
'
'                strData = Data(iRow, iCol)
'
'                Select Case iCol
'                Case 0 '구분
'                    Sheets("가격수령").Cells(36, 3).Value = "A" + strData
'                Case 1 '번호
'                    Sheets("가격수령").Cells(36, 4).Value = strData
'                Case 2 '계좌번호
'                    Sheets("가격수령").Cells(36, 5).Value = strData
'                Case 3 '종목명
'                    Sheets("가격수령").Cells(36, 7).Value = strData
'                Case 4 '체결가
'                    Sheets("가격수령").Cells(36, 8).Value = strData
'                Case 5 '체결수량
'                    Sheets("가격수령").Cells(36, 9).Value = strData
'                Case 6 '종목코드
'                    Sheets("가격수령").Cells(36, 6).Value = strData
'                End Select
'            End If
            
'        Next
'    Next

End Sub

Private Sub OrderObject_ReceiveQuote(ByVal bsSheetName As String, ByVal nCellY As Long, ByVal nCellX As Long, ByVal bsFid As String, ByVal bsData As String)
'변수 선언 부분

'On Error Resume Next
On Error GoTo ErrMsg

'    DoEvents
    
    With Sheets(Int(bsSheetName))
        Select Case Int(bsFid)
        Case 15001 '현재가
            .Cells(nCellY + 0, nCellX).Value = bsData
'        Case 15006 '등락구분 '현재 아무처리 안함
'            .Cells(nCellY + 1, nCellX).Value = bsData
        Case 14531 '매수호가1
            .Cells(nCellY + 1, nCellX).Value = bsData
        Case 14541 '매수호가잔량1
            .Cells(nCellY + 2, nCellX).Value = bsData
        Case 14501 '매도호가1
            .Cells(nCellY + 3, nCellX).Value = bsData
        Case 14511 '매도호가잔량1
            .Cells(nCellY + 4, nCellX).Value = bsData
        Case 16188 '선물이론가
            .Cells(nCellY + 5, nCellX).Value = bsData
        Case 16175 '최종거래일
            .Cells(nCellY + 6, nCellX).Value = Mid(bsData, 1, 4) + "-" + Mid(bsData, 5, 2) + "-" + Mid(bsData, 7, 2)
        End Select
    End With
    
    Exit Sub
ErrMsg:
    If (Err.Number = 50290) Then Exit Sub
    If (Err.Number <> 0) And (Err.Number <> 50290) Then
        MsgBox "ReceiveQuote 오류번호 : " & Err.Number & vbCr & _
       "오류내용 : " & Err.Description, vbCritical, "오류"
    End If
    Exit Sub

End Sub

Private Sub OrderObject_ReceiveQuoteF(ByVal bsSheetName As String, ByVal nCellY As Long, ByVal nCellX As Long, ByVal bsFid As String, ByVal bsData As String)
'변수 선언 부분

'On Error Resume Next
On Error GoTo ErrMsg

'    DoEvents
    
    With Sheets(Int(bsSheetName))
        Select Case Int(bsFid)
        Case 15001 '현재가
            .Cells(nCellY + 0, nCellX).Value = bsData
'        Case 15006 '등락구분 '현재 아무처리 안함
'            .Cells(nCellY + 1, nCellX).Value = bsData
        Case 14531 '매수호가1
            .Cells(nCellY + 1, nCellX).Value = bsData
        Case 14541 '매수호가잔량1
            .Cells(nCellY + 2, nCellX).Value = bsData
        Case 14501 '매도호가1
            .Cells(nCellY + 3, nCellX).Value = bsData
        Case 14511 '매도호가잔량1
            .Cells(nCellY + 4, nCellX).Value = bsData
        Case 16188 '선물이론가
            .Cells(nCellY + 5, nCellX).Value = bsData
        Case 16175 '최종거래일
            If bsData <> "0" Then
                .Cells(nCellY + 6, nCellX).Value = Mid(bsData, 1, 4) + "-" + Mid(bsData, 5, 2) + "-" + Mid(bsData, 7, 2)
            Else
                .Cells(nCellY + 6, nCellX).Value = ""
            End If
        End Select
    End With
    
    Exit Sub
ErrMsg:
    If (Err.Number = 50290) Then Exit Sub
    If (Err.Number <> 0) And (Err.Number <> 50290) Then
        MsgBox "ReceiveQuote 오류번호 : " & Err.Number & vbCr & _
       "오류내용 : " & Err.Description, vbCritical, "오류"
    End If
    Exit Sub

End Sub

Private Sub OrderObject_ReceiveResult(ByVal bsData As String)

End Sub

Private Sub OrderObject_ReceiveResult2(ByVal bsData As String, ByVal nRowSize As Long, ByVal nColSize As Long)
    '조회
    Dim data() As String
    Dim iRow As Integer
    Dim iCol As Integer
    Dim c As Integer
    
    Dim strData As String
    
    If (nColSize < 0) Or (nRowSize < 0) Then
        Exit Sub
    End If
    
'    Application.ScreenUpdating = False
    'Application.DisplayStatusBar = False
'    Application.Calculation = xlCalculationManual
    'Application.EnableEvents = False
    'ActiveSheet.DisplayPageBreaks = False
    
    
    If StrComp(bsData, "HONAQ001") = 0 Then '잔고조회
        ReDim HONAQ001Data(nRowSize, nColSize) As String
        HONAQ001Count = nRowSize
        Sheets("모듈").OrderObject.GetRepeatResult HONAQ001Data
        
        Application.Calculation = xlCalculationManual
        
        For c = QuoteStockHStartPos To QuoteStockItemCount + QuoteStockHStartPos - 1
            Sheets("가격수령").Cells(15, c).Value = 0
            Sheets("가격수령").Cells(16, c).Value = 0
        Next
        
        For iRow = 0 To UBound(HONAQ001Data, 1) - 1
            For c = QuoteStockHStartPos To QuoteStockItemCount + QuoteStockHStartPos - 1
                If (Sheets("가격수령").Cells(3, c).Value = "A" + HONAQ001Data(iRow, 0)) Then
                    Sheets("가격수령").Cells(15, c).Value = HONAQ001Data(iRow, 3)
                    Sheets("가격수령").Cells(16, c).Value = HONAQ001Data(iRow, 6)
                End If
            Next
                
'            Sheets("가격수령").Range("C3:GT3").Find (HONAQ001Data(iRow, 0))
'            Select Case iCol
'            Case 0 '종목코드(표준단축종목코드로 변경)
'                Sheets("가격수령").Cells(iRow + 7, 2).Value = "A" + strData
'            Case 1 '종목명
'                Sheets("가격수령").Cells(iRow + 7, 1).Value = strData
'            Case 3 '보유수량
'                Sheets("가격수령").Cells(iRow + 7, 3).Value = strData
'            Case 5 '주문가능수량
'                Sheets("가격수령").Cells(iRow + 7, 4).Value = strData
'            Case 6 '매수평균가
'                Sheets("가격수령").Cells(iRow + 7, 5).Value = strData
'            Case 7 '매수금액
'                Sheets("가격수령").Cells(iRow + 7, 6).Value = strData
'            Case 12 '현재가
'                Sheets("가격수령").Cells(iRow + 7, 7).Value = strData
'            Case 14 '잔고평가금액
'                Sheets("가격수령").Cells(iRow + 7, 8).Value = strData
'            Case 15 '잔고평가손익
'                Sheets("가격수령").Cells(iRow + 7, 9).Value = strData
'            Case 16 '잔고평가손익률
'                Sheets("가격수령").Cells(iRow + 7, 10).Value = strData
'            End Select
        Next
        Application.Calculation = xlCalculationAutomatic

    ElseIf StrComp(bsData, "HONBQ001") = 0 Then '선물잔고조회
        ReDim HONBQ001Data(nRowSize, nColSize) As String
        HONBQ001Count = nRowSize
        Sheets("모듈").OrderObject.GetRepeatResult HONBQ001Data
        
        Application.Calculation = xlCalculationManual
        
        For c = QuoteFutureHStartPos To QuoteFutureItemCount + QuoteFutureHStartPos - 1
            Sheets("가격수령").Cells(15, c).Value = 0
            Sheets("가격수령").Cells(16, c).Value = 0
        Next
        
        For iRow = 0 To UBound(HONBQ001Data, 1) - 1
            For c = QuoteFutureHStartPos To QuoteFutureItemCount + QuoteFutureHStartPos - 1
                If (Sheets("가격수령").Cells(3, c).Value = HONBQ001Data(iRow, 0)) Then
                    Sheets("가격수령").Cells(15, c).Value = HONBQ001Data(iRow, 5)
                    Sheets("가격수령").Cells(16, c).Value = HONBQ001Data(iRow, 7)
                End If
            Next
            
'            For iCol = 0 To UBound(HONBQ001Data, 2)
'                strData = HONBQ001Data(iRow, iCol)
'                Sheets("가격수령").Activate
'                Select Case iCol
'                Case 0 '종목코드
'                    Sheets("가격수령").Cells(iRow + 7, 2).Value = strData
'                Case 1 '종목명
'                    Sheets("가격수령").Cells(iRow + 7, 1).Value = strData
'                Case 3 '매매구분(매도,매수)
'                    Sheets("가격수령").Cells(iRow + 7, 3).Value = strData
'                Case 5 '보유수량
'                    Sheets("가격수령").Cells(iRow + 7, 4).Value = strData
'                Case 6 '청산가능수량
'                    Sheets("가격수령").Cells(iRow + 7, 5).Value = strData
'                Case 7 '평균단가
'                    Sheets("가격수령").Cells(iRow + 7, 6).Value = strData
'                Case 9 '현재가
'                    Sheets("가격수령").Cells(iRow + 7, 7).Value = strData
'                Case 11 '평가금액
'                    Sheets("가격수령").Cells(iRow + 7, 8).Value = strData
'                Case 12 '평가손익
'                    Sheets("가격수령").Cells(iRow + 7, 9).Value = strData
'                Case 13 '평가수익률
'                    Sheets("가격수령").Cells(iRow + 7, 10).Value = strData
'                End Select
'
'            Next
        Next
        Application.Calculation = xlCalculationAutomatic
        
    ElseIf StrComp(bsData, "HONAQ003") = 0 Then '체결조회
        On Error Resume Next
        Application.CutCopyMode = False
        On Error GoTo 0 ' 에러 처리를 다시 활성화합니다.
        DoEvents
        WaitFor (1)
        ReDim HONAQ003Data(nRowSize, nColSize) As String
        Sheets("모듈").OrderObject.GetRepeatResult HONAQ003Data
        If isSelectHONAQ003 Then
            Application.Calculation = xlCalculationManual
            For iRow = 0 To UBound(HONAQ003Data, 1) - 1
                For iCol = 0 To UBound(HONAQ003Data, 2)
                    strData = HONAQ003Data(iRow, iCol)
                    Select Case iCol
                    Case 0 '주문번호
                        Sheets("체결내역").Cells(iRow + 3, 1).Value = strData
                    Case 1 '원주문번호
                        Sheets("체결내역").Cells(iRow + 3, 2).Value = strData
                    Case 2 '주문시간
                        Sheets("체결내역").Cells(iRow + 3, 3).Value = strData
                    Case 3 '종목코드
                        Sheets("체결내역").Cells(iRow + 3, 5).Value = "A" + strData
                    Case 4 '종목명
                        Sheets("체결내역").Cells(iRow + 3, 4).Value = strData
                    Case 9 '주문구분명
                        Sheets("체결내역").Cells(iRow + 3, 6).Value = strData
                    Case 11 '주문수량
                        Sheets("체결내역").Cells(iRow + 3, 7).Value = strData
                    Case 12 '주문가
                        Sheets("체결내역").Cells(iRow + 3, 8).Value = strData
                    Case 13 '체결수량
                        Sheets("체결내역").Cells(iRow + 3, 9).Value = strData
                    Case 14 '체결가
                        Sheets("체결내역").Cells(iRow + 3, 10).Value = strData
                    Case 15 '미체결수량
                        Sheets("체결내역").Cells(iRow + 3, 11).Value = strData
                    Case 21 '체결상태명
                        Sheets("체결내역").Cells(iRow + 3, 12).Value = strData
                    End Select
                Next
            Next
            Application.Calculation = xlCalculationAutomatic
        Else
            HONAQ003Count = nRowSize
        End If
        
'        For iRow = 0 To UBound(HONAQ003Data, 1) - 1
'            For iCol = 0 To UBound(HONAQ003Data, 2)
'                strData = HONAQ003Data(iRow, iCol)
'                Select Case iCol
'                Case 0 '주문번호
'                    Sheets("가격수령").Cells(iRow + 21, 1).Value = strData
'                Case 1 '원주문번호
'                    Sheets("가격수령").Cells(iRow + 21, 2).Value = strData
'                Case 2 '주문시간
'                    Sheets("가격수령").Cells(iRow + 21, 3).Value = strData
'                Case 3 '종목코드
'                    Sheets("가격수령").Cells(iRow + 21, 5).Value = "A" + strData
'                Case 4 '종목명
'                    Sheets("가격수령").Cells(iRow + 21, 4).Value = strData
'                Case 9 '주문구분명
'                    Sheets("가격수령").Cells(iRow + 21, 6).Value = strData
'                Case 11 '주문수량
'                    Sheets("가격수령").Cells(iRow + 21, 7).Value = strData
'                Case 12 '주문가
'                    Sheets("가격수령").Cells(iRow + 21, 8).Value = strData
'                Case 13 '체결수량
'                    Sheets("가격수령").Cells(iRow + 21, 9).Value = strData
'                Case 14 '체결가
'                    Sheets("가격수령").Cells(iRow + 21, 10).Value = strData
'                Case 15 '미체결수량
'                    Sheets("가격수령").Cells(iRow + 21, 11).Value = strData
'                Case 21 '체결상태명
'                    Sheets("기본주식주문").Cells(iRow + 21, 12).Value = strData
'                End Select
'            Next
'        Next
    ElseIf StrComp(bsData, "HONBQ003") = 0 Then '선물체결조회
        ReDim HONBQ003Data(nRowSize, nColSize) As String
        HONBQ003Count = nRowSize
        Sheets("모듈").OrderObject.GetRepeatResult HONBQ003Data

'        For iRow = 0 To UBound(HONBQ003Data, 1) - 1
'            For iCol = 0 To UBound(HONBQ003Data, 2)
'                strData = HONBQ003Data(iRow, iCol)
'                Select Case iCol
'                Case 1 '주문번호
'                    Sheets("가격수령").Cells(iRow + 21, 1).Value = strData
'                Case 2 '원주문번호
'                    Sheets("가격수령").Cells(iRow + 21, 2).Value = strData
'                Case 3 '주문시간
'                    Sheets("가격수령").Cells(iRow + 21, 3).Value = strData
'                Case 4 '종목코드
'                    Sheets("가격수령").Cells(iRow + 21, 5).Value = strData
'                Case 5 '종목명
'                    Sheets("가격수령").Cells(iRow + 21, 4).Value = strData
'                Case 13 '주문구분명
'                    Sheets("가격수령").Cells(iRow + 21, 6).Value = strData
'                Case 12 '주문수량
'                    Sheets("가격수령").Cells(iRow + 21, 7).Value = strData
'                Case 11 '주문가
'                    Sheets("가격수령").Cells(iRow + 21, 8).Value = strData
'                Case 16 '체결수량
'                    Sheets("가격수령").Cells(iRow + 21, 9).Value = strData
'                Case 15 '체결가
'                    Sheets("가격수령").Cells(iRow + 21, 10).Value = strData
'                Case 21 '미체결수량
'                    Sheets("가격수령").Cells(iRow + 21, 11).Value = strData
'                Case 14 '체결상태명
'                    Sheets("가격수령").Cells(iRow + 21, 12).Value = strData
'                End Select
'            Next
'        Next
    ElseIf StrComp(bsData, "HONAQ001_H") = 0 Then '주식잔고조회헤더
'        ReDim Data(nRowSize, nColSize) As String
'        Sheets("모듈").OrderObject.GetRepeatHResult Data
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'
'                Sheets("주식잔고").Activate
'                Select Case iCol
'                Case 0
'                    Sheets("주식잔고").Cells(3, 2).Value = strData
'                Case 1
'                    Sheets("주식잔고").Cells(4, 2).Value = strData
'                Case 2
'                    Sheets("주식잔고").Cells(3, 6).Value = strData
'                Case 3
'                    Sheets("주식잔고").Cells(4, 6).Value = strData
'                Case 4
'                    Sheets("주식잔고").Cells(4, 9).Value = strData
'                Case 5
'                    Sheets("주식잔고").Cells(4, 10).Value = strData
'                Case 6
'                    Sheets("주식잔고").Cells(3, 9).Value = strData
'                Case 7
'                    Sheets("주식잔고").Cells(3, 10).Value = strData
'                End Select
'            Next
'        Next
    ElseIf StrComp(bsData, "HONBQ001_H") = 0 Then '선옵잔고조회헤더
'        ReDim Data(nRowSize, nColSize) As String
'        Sheets("모듈").OrderObject.GetRepeatHResult Data
'        For iRow = 0 To UBound(Data, 1) - 1
'            For iCol = 0 To UBound(Data, 2)
'                strData = Data(iRow, iCol)
'
'                Sheets("선옵잔고").Activate
'                Select Case iCol
'                Case 0
'                    Sheets("선옵잔고").Cells(3, 2).Value = strData
'                Case 1
'                    Sheets("선옵잔고").Cells(4, 2).Value = strData
'                Case 2
'                    Sheets("선옵잔고").Cells(3, 6).Value = strData
'                Case 3
'                    Sheets("선옵잔고").Cells(4, 6).Value = strData
'                Case 4
'                    Sheets("선옵잔고").Cells(3, 9).Value = strData
'                Case 5
'                    Sheets("선옵잔고").Cells(4, 9).Value = strData
'                End Select
'            Next
'        Next
    End If
    
'    Application.ScreenUpdating = screenUpdateState
    'Application.DisplayStatusBar = statusBarState
    'Application.Calculation = calcState
    'Application.EnableEvents = eventsState
    'ActiveSheet.DisplayPageBreaks = displayPageBreakState
End Sub

Public Sub QuoteCallFutures_Click()
'    Dim sIndex As Integer
'    Dim sItemCodes As String
'    Dim iCol As Integer
'
'    sIndex = Sheets("가격수령").Index
'
'    'Application.ScreenUpdating = False
'    'Application.DisplayStatusBar = False
'    Application.Calculation = xlCalculationManual
'    'Application.EnableEvents = False
'    'ActiveSheet.DisplayPageBreaks = False
'
'    sItemCodes = ""
'
'    For iCol = 1 To QuoteFutureItemCount
'        sItemCodes = sItemCodes & Sheets("가격수령").Cells(QuoteFutureVStartPos, iCol + QuoteFutureHStartPos - 1).Value & ","
'        If iCol = QuoteFutureItemCount Then
'            sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
'        End If
'    Next
'    'MsgBox sItemCodes
'
'    With Sheets("모듈")
'        'sItemCodes = "101Q6000,101Q9000,101QC000,101R3000,101R6000,101R9000,101RC000,101S3000"
'        .OrderObject.SetQuote sItemCodes, sIndex, 6, QuoteFutureHStartPos, 0, 5, 0
'    End With


'20210127-신규
    '호출 하기 전에 한번만.
    Application.Calculation = xlCalculationManual
    
    
'CallQuote Parameters
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
    
    '선물시세 호출
    CallQuote 2, "가격수령", "모듈", Sheets("가격수령").Index, QuoteFutureVStartPos, QuoteFutureHStartPos, 6, 404, QuoteFutureItemCount, QuoteFutureItemCount, 0, 5, 0, 0.1

    
'    '개별시세 호출
'    CallQuote 2, "가격수령", "모듈", Sheets("가격수령").Index, 3, 421, 6, 421, 1530, 50, 0, 0, 0, 0.1
''    CallQuote 2, "가격수령", "모듈", Sheets("가격수령").Index, 3, 421, 6, 421, 100, 50

    Sheets("모듈").OrderObject.CallQuoteComplete 2

End Sub

Public Sub QuoteCallStock_Click()
'    Dim sIndex As Integer
'    Dim sItemCodes As String
'    Dim sOrgCode As String
'    Dim iCol As Integer
'    sIndex = Sheets("가격수령").Index
'
'    'Application.ScreenUpdating = False
'    'Application.DisplayStatusBar = False
'    Application.Calculation = xlCalculationManual
'    'Application.EnableEvents = False
'    'ActiveSheet.DisplayPageBreaks = False
'
'    With Sheets("모듈")
'        sItemCodes = ""
'        For iCol = 1 To QuoteStockItemCount
'            sOrgCode = Sheets("가격수령").Cells(QuoteStockVStartPos, iCol + QuoteStockHStartPos - 1).Value
'            sItemCodes = sItemCodes & Right(sOrgCode, Len(sOrgCode) - 1) & ","
'            If iCol Mod 50 = 0 Then
'                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
'                'MsgBox sItemCodes
'                Sheets("모듈").OrderObject.SetQuote sItemCodes, sIndex, 6, iCol - 50 + 3
'                sItemCodes = ""
'                WaitFor (0.1)
'            End If
'        Next
'
'        sItemCodes = "51"
'        Sheets("모듈").OrderObject.SetQuote sItemCodes, sIndex, 5, QuoteFutureHStartPos - 1
'    End With


'20210127-신규
    '호출 하기 전에 한번만.
    Application.Calculation = xlCalculationManual

'CallQuote Parameters
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
    
    '주식 시세 호출
    CallQuote 1, "가격수령", "모듈", Sheets("가격수령").Index, QuoteStockVStartPos, QuoteStockHStartPos, 6, 3, QuoteStockItemCount, 50, 0, 0, 0, 0.1


'SetQuote Parameters
'    BSTR sItemCode                '종목리스트
'    INT nSheetNum                 '종목을 받는 엑셀 시트의 인덱스 번호
'    INT nCellY                    '세로 시세 시작 위치
'    INT nCellX                    '가로 시세 시작 위치
'    INT nDirection                '방향 : 0(X축), 1(Y축)-종목이 X축 방향, Y축 방향으로 위치해 있느냐를 선택하는 옵션
'    INT nTimer                    '타이머 : n(n초에 한번씩 시세를 뿌려줌)
'    INT nTimerType                '타이머구분 : 0(시세를 호출한 시점 부터 시작), 1(정시를 기준으로 처리-시작시간이 3초에 시작된경우 5초에 처리됨)

    
    '코스피 200 지수 호출
    Sheets("모듈").OrderObject.SetQuote "51", Sheets("가격수령").Index, 5, QuoteFutureHStartPos - 1
    
    Sheets("모듈").OrderObject.CallQuoteComplete 1
    
End Sub

Public Sub QuotePause_Click()
'    Dim sItemCodes As String
'    Dim sOrgCode As String
'    Dim iCol As Integer
'
'
'    With Sheets("모듈")
'        sItemCodes = ""
'        For iCol = 1 To QuoteStockItemCount
'            sOrgCode = Sheets("가격수령").Cells(QuoteStockVStartPos, iCol + QuoteStockHStartPos - 1).Value
'            sItemCodes = sItemCodes & Right(sOrgCode, Len(sOrgCode) - 1) & ","
'            If iCol Mod 50 = 0 Then
'                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
'                'MsgBox sItemCodes
'                Sheets("모듈").OrderObject.StopQuote sItemCodes
'                sItemCodes = ""
'                WaitFor (0.1)
'            End If
'        Next
'
'        sItemCodes = "51"
'        Sheets("모듈").OrderObject.StopQuote sItemCodes
'
'        sItemCodes = ""
'
'        For iCol = 1 To QuoteFutureItemCount
'            sItemCodes = sItemCodes & Sheets("가격수령").Cells(QuoteFutureVStartPos, iCol + QuoteFutureHStartPos - 1).Value & ","
'            If iCol = QuoteFutureItemCount Then
'                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
'            End If
'        Next
'
'        .OrderObject.StopQuote sItemCodes
'    End With
'
'    Application.Calculation = xlCalculationAutomatic


'20210127-신규
'PauseQuote Parameters
'    pnMarketType As String       '시장 구분(1:장내종목(코스피, 코스닥), 2:선물옵션)
'    psItemSheetName As String    '종목 리스트가 있는 엑셀 시트 명
'    psModuleSheetName As String  '모듈이 있는 엑셀 시트 명
'    pnItemY As Integer           '세로 종목 위치
'    pnItemX As Integer           '가로 종목 위치
'    pnTotalItemCount As Integer  '총 종목수
'    pnCallMaxCount As Integer    '분할 호출 개수 : 최대 50종목까지만 호출 할 수 있음
'    pnWaitTime As Single         '시세 호출 후 대기 시간


    '주식 시세 일시중지
    PauseQuote 1, "가격수령", "모듈", QuoteStockVStartPos, QuoteStockHStartPos, QuoteStockItemCount, 50, 0.1

'StopQuote Parameters
'    BSTR sItemCode                '종목리스트

    '코스피 200 지수 일시중지
    Sheets("모듈").OrderObject.StopQuote "51"

    '선물 시세 일시중지
    PauseQuote 2, "가격수령", "모듈", QuoteFutureVStartPos, QuoteFutureHStartPos, QuoteFutureItemCount, QuoteFutureItemCount, 0.1

'    '개별 시세 일시중지
'    PauseQuote 2, "가격수령", "모듈", 3, 421, 1530, 50, 0.1
''    PauseQuote 2, "가격수령", "모듈", 3, 421, 100, 50

    '다 호출 하고 나서 자동으로 변경 해야됨.
    Application.Calculation = xlCalculationAutomatic


End Sub

Public Sub QuoteStart_Click()
'    Dim sItemCodes As String
'    Dim sOrgCode As String
'    Dim iCol As Integer
'
'    Application.Calculation = xlCalculationManual
'
'    With Sheets("모듈")
'        sItemCodes = ""
'        For iCol = 1 To QuoteStockItemCount
'            sOrgCode = Sheets("가격수령").Cells(QuoteStockVStartPos, iCol + QuoteStockHStartPos - 1).Value
'            sItemCodes = sItemCodes & Right(sOrgCode, Len(sOrgCode) - 1) & ","
'            If iCol Mod 50 = 0 Then
'                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
'                'MsgBox sItemCodes
'                Sheets("모듈").OrderObject.StartQuote sItemCodes
'                sItemCodes = ""
'                WaitFor (0.1)
'            End If
'        Next
'
'        sItemCodes = "51"
'        Sheets("모듈").OrderObject.StartQuote sItemCodes
'
'        sItemCodes = ""
'
'        For iCol = 1 To QuoteFutureItemCount
'            sItemCodes = sItemCodes & Sheets("가격수령").Cells(QuoteFutureVStartPos, iCol + QuoteFutureHStartPos - 1).Value & ","
'            If iCol = QuoteFutureItemCount Then
'                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
'            End If
'        Next
'        .OrderObject.StartQuote sItemCodes
'    End With


'20210127-신규
'StartQuote Parameters
'    pnMarketType As String       '시장 구분(1:장내종목(코스피, 코스닥), 2:선물옵션)
'    psItemSheetName As String    '종목 리스트가 있는 엑셀 시트 명
'    psModuleSheetName As String  '모듈이 있는 엑셀 시트 명
'    pnItemY As Integer           '세로 종목 위치
'    pnItemX As Integer           '가로 종목 위치
'    pnTotalItemCount As Integer  '총 종목수
'    pnCallMaxCount As Integer    '분할 호출 개수 : 최대 50종목까지만 호출 할 수 있음
'    pnWaitTime As Single         '시세 호출 후 대기 시간


    '호출 하기 전에 한번만.
    Application.Calculation = xlCalculationManual

    '주식 시세 일시중지
    StartQuote 1, "가격수령", "모듈", QuoteStockVStartPos, QuoteStockHStartPos, QuoteStockItemCount, 50, 0.1

'StopQuote Parameters
'    BSTR sItemCode                '종목리스트

    '코스피 200 지수 일시중지
    Sheets("모듈").OrderObject.StartQuote "51"

    '선물 시세 일시중지
    StartQuote 2, "가격수령", "모듈", QuoteFutureVStartPos, QuoteFutureHStartPos, QuoteFutureItemCount, QuoteFutureItemCount, 0.1

'    '개별 시세 일시중지
'    StartQuote 2, "가격수령", "모듈", 3, 421, 1530, 50, 0.1
''    StartQuote 2, "가격수령", "모듈", 3, 421, 100, 50


End Sub

Public Sub QuoteStop_Click()
'    Dim sItemCodes As String
'    Dim sOrgCode As String
'    Dim iCol As Integer
'
'
'    With Sheets("모듈")
'        sItemCodes = ""
'        For iCol = 1 To QuoteStockItemCount
'            sOrgCode = Sheets("가격수령").Cells(QuoteStockVStartPos, iCol + QuoteStockHStartPos - 1).Value
'            sItemCodes = sItemCodes & Right(sOrgCode, Len(sOrgCode) - 1) & ","
'            If iCol Mod 50 = 0 Then
'                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
'                'MsgBox sItemCodes
'                Sheets("모듈").OrderObject.ResetQuote sItemCodes
'                sItemCodes = ""
'                WaitFor (0.1)
'            End If
'        Next
'
'        sItemCodes = "51"
'        Sheets("모듈").OrderObject.ResetQuote sItemCodes
'
'        sItemCodes = ""
'
'        For iCol = 1 To QuoteFutureItemCount
'            sItemCodes = sItemCodes & Sheets("가격수령").Cells(QuoteFutureVStartPos, iCol + QuoteFutureHStartPos - 1).Value & ","
'            If iCol = QuoteFutureItemCount Then
'                sItemCodes = Left(sItemCodes, Len(sItemCodes) - 1)
'            End If
'        Next
'
'        .OrderObject.ResetQuote sItemCodes
'    End With
'
'    Application.Calculation = xlCalculationAutomatic


'20210127-신규
'StopQuote Parameters
'    pnMarketType As String       '시장 구분(1:장내종목(코스피, 코스닥), 2:선물옵션)
'    psItemSheetName As String    '종목 리스트가 있는 엑셀 시트 명
'    psModuleSheetName As String  '모듈이 있는 엑셀 시트 명
'    pnItemY As Integer           '세로 종목 위치
'    pnItemX As Integer           '가로 종목 위치
'    pnTotalItemCount As Integer  '총 종목수
'    pnCallMaxCount As Integer    '분할 호출 개수 : 최대 50종목까지만 호출 할 수 있음
'    pnWaitTime As Single         '시세 호출 후 대기 시간

    '주식 시세 일시중지
    StopQuote 1, "가격수령", "모듈", QuoteStockVStartPos, QuoteStockHStartPos, QuoteStockItemCount, 50, 0.1

'ResetQuote Parameters
'    BSTR sItemCode                '종목리스트

    '코스피 200 지수 일시중지
    Sheets("모듈").OrderObject.ResetQuote "51"

    '선물 시세 일시중지
    StopQuote 2, "가격수령", "모듈", QuoteFutureVStartPos, QuoteFutureHStartPos, QuoteFutureItemCount, QuoteFutureItemCount, 0.1

'    '개별 시세 일시중지
'    StopQuote 2, "가격수령", "모듈", 3, 421, 1530, 50, 0.1
''    StopQuote 2, "가격수령", "모듈", 3, 421, 100, 50

    '다 호출 하고 나서 자동으로 변경 해야됨.
    Application.Calculation = xlCalculationAutomatic


End Sub

Public Sub StockBatchBuy_Click()
    Dim sIssueCode As String    '코드
    Dim sAccount As String    '계좌번호
    Dim sPassword As String    '비밀번호
    Dim sOrderType As String    '주문조건
    Dim sCallType As String    '호가유형
    Dim sQty As String    '매매수량
    Dim sPrice As String    '매매가격
    Dim i As Integer

    AccountPass = Sheets("모듈").Cells(3, 2).Value

    sAccount = AccountList(0, 0)
    sPassword = AccountPass
    
    sOrderType = "1" '"1" 일반, "2" IOC, "3" FOK
    sCallType = "00" '"00" 지정가, "03" 시장가, "05" 조건부, "06" 최유리, "07" 최우선
    
    'CreateBizObject
    
    For i = 1 To QuoteStockItemCount
        sIssueCode = Sheets("가격수령").Cells(3, i + QuoteStockHStartPos - 1).Value
        sPrice = Sheets("가격수령").Cells(29, i + QuoteStockHStartPos - 1).Value
        sQty = Sheets("가격수령").Cells(30, i + QuoteStockHStartPos - 1).Value
        If (sPrice <> "0") And (sQty <> "0") Then
            With Sheets("모듈")
                .OrderObject.SendBuy sAccount, sPassword, sIssueCode, sOrderType, sCallType, sQty, sPrice
                WaitFor (0.01)
            End With
        End If
        
    Next

End Sub

Public Sub StockBatchCancel_Click()
    Dim sIssueCode As String    '코드
    Dim sAccount As String    '계좌번호
    Dim sPassword As String    '비밀번호
    Dim sOrgOrdNum As String    '원주문번호
    Dim sQty As String    '매수수량
    Dim iRow As Integer
    Dim iCol As Integer
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value

    sAccount = AccountList(0, 0)
    sPassword = AccountPass

    'CreateBizObject
    
    
    For iRow = 0 To UBound(HONAQ003Data, 1) - 1
        sIssueCode = HONAQ003Data(iRow, 3)
        sOrgOrdNum = HONAQ003Data(iRow, 0)
        sQty = HONAQ003Data(iRow, 15)
        With Sheets("모듈")
            .OrderObject.SendCancel sAccount, sPassword, sIssueCode, sOrgOrdNum, sQty
            WaitFor (0.01)
        End With
    Next
End Sub

Public Sub StockBatchSell_Click()
    Dim sIssueCode As String    '코드
    Dim sAccount As String    '계좌번호
    Dim sPassword As String    '비밀번호
    Dim sOrderType As String    '주문조건
    Dim sCallType As String    '호가유형
    Dim sQty As String    '매매수량
    Dim sPrice As String    '매매가격
    Dim i As Integer
    
    AccountPass = Sheets("모듈").Cells(3, 2).Value

    sAccount = AccountList(0, 0)
    sPassword = AccountPass

    sOrderType = "1" '"1" 일반, "2" IOC, "3" FOK
    sCallType = "00" '"00" 지정가, "03" 시장가, "05" 조건부, "06" 최유리, "07" 최우선

    
    'CreateBizObject
    
    For i = 1 To QuoteStockItemCount
        sIssueCode = Sheets("가격수령").Cells(3, i + QuoteStockHStartPos - 1).Value
        sPrice = Sheets("가격수령").Cells(31, i + QuoteStockHStartPos - 1).Value
        sQty = Sheets("가격수령").Cells(32, i + QuoteStockHStartPos - 1).Value
        If (sPrice <> "0") And (sQty <> "0") Then
            With Sheets("모듈")
                .OrderObject.SendSell sAccount, sPassword, sIssueCode, sOrderType, sCallType, sQty, sPrice
                WaitFor (0.01)
            End With
        End If
            
    Next

End Sub

Private Sub WapiCommunityMngCtrl_NormalRealReceive(ByVal pRealObj As Long, ByVal sPacketName As String, ByVal sPrimaryKey As String)

End Sub

Private Sub WapiCommunityMngCtrl_OrderRealReceive(ByVal pRealObj As Long, ByVal sPacketName As String, ByVal sPrimaryKey As String)

'    Dim strPrice As String
'    Dim strMsg As String
'
'    Sheets("모듈").Cells(23, 11).Value = CStr(pRealObj) + ", " + sPacketName + ", " + sPrimaryKey
'    If sPacketName = "SC0" Then
'    '101D3000
'        strPrice = WapiCommunityMngCtrl.GetRealItem(pRealObj, 3)
'        Sheets("모듈").Cells(23, 12).Value = strPrice
'
'    ElseIf sPacketName = "XF1" Then
'        strMsg = strMsg + "유저아이디: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 0) + ", "
'        strMsg = strMsg + "계좌번호: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 1) + ", "
'        strMsg = strMsg + "주문번호: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 2) + ", "
'        strMsg = strMsg + "원주문번호: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 3) + ", "
'        strMsg = strMsg + "종목코드: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 4) + ", "
'        strMsg = strMsg + "현재가: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 15)
'
'    ElseIf sPacketName = "XX1" Then
'        strMsg = strMsg + "유저아이디: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 0) + ", "
'        strMsg = strMsg + "계좌번호: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 1) + ", "
'        strMsg = strMsg + "주문번호: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 2) + ", "
'        strMsg = strMsg + "원주문번호: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 3) + ", "
'        strMsg = strMsg + "종목코드: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 4) + ", "
'        strMsg = strMsg + "현재가: " + WapiCommunityMngCtrl.GetRealItem(pRealObj, 6)
'
'    End If
'
'    Sheets("모듈").Cells(23, 13) = strMsg
    
End Sub

Private Sub WapiCommunityMngCtrl_TrReceive(ByVal sTrNum As String, ByVal strData As Long)
'' TR응답
'    Dim nCount1 As Integer
'    Dim strCount1 As String
'    Dim strKeyCount1 As String
'    Dim strKeyName1 As String
'    Dim i, iRow, c As Integer
'
'    'MsgBox sTrNum
'
'    Select Case sTrNum
'    Case "09004": '계좌리스트 조회
'
'        strCount1 = Trim(WapiCommunityMngCtrl.GetTrData("09004.데이터카운트"))
'        If strCount1 <> "" Then
'            nCount1 = CInt(strCount1)
'        End If
'
'        For i = 0 To nCount1 - 1
'            strKeyCount1 = CStr(i)
'
'            strKeyName1 = "09004.계좌번호종류." + strKeyCount1
'            Sheets("모듈").Cells(9 + i, 13).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'
'            strKeyName1 = "09004.계좌번호." + strKeyCount1
'            Sheets("모듈").Cells(9 + i, 14).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'
'            strKeyName1 = "09004.계좌명." + strKeyCount1
'            Sheets("모듈").Cells(9 + i, 15).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'
'            strKeyName1 = "09004.계좌비밀번호." + strKeyCount1
'            Sheets("모듈").Cells(9 + i, 16).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'
'            strKeyName1 = "09004.대표계좌번호." + strKeyCount1
'            Sheets("모듈").Cells(9 + i, 17).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'        Next i
'
'    Case "96062": '주문결과
'
'        'Sheets("모듈").Cells(8, 9).Value = "주문번호: " + Trim(WapiCommunityMngCtrl.GetTrData("96062.주문번호")) + ", 계좌번호: " + Trim(WapiCommunityMngCtrl.GetTrData("96062.계좌번호")) + ", 종목코드: " + Trim(WapiCommunityMngCtrl.GetTrData("96062.종목코드"))
'
'    Case "96007": '주식:잔고조회
'
'        strCount1 = Trim(WapiCommunityMngCtrl.GetTrData("96007.데이터카운트"))
'        If strCount1 <> "" Then
'            nCount1 = CInt(strCount1)
'        End If
'
'        ReDim HONAQ001Data(nCount1, 21) As String
'
'        HONAQ001Count = nCount1
'
'        Application.Calculation = xlCalculationManual
'
'        For i = 0 To nCount1 - 1
'            strKeyCount1 = CStr(i)
'
'            strKeyName1 = "96007.신용구분." + strKeyCount1
'            HONAQ001Data(i, 0) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.시장구분." + strKeyCount1
'            HONAQ001Data(i, 1) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.종목명." + strKeyCount1
'            HONAQ001Data(i, 2) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.매도가능수량." + strKeyCount1
'            HONAQ001Data(i, 3) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.잔고수량." + strKeyCount1
'            HONAQ001Data(i, 4) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.전일대비." + strKeyCount1
'            HONAQ001Data(i, 5) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.등락율." + strKeyCount1
'            HONAQ001Data(i, 6) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.평가손익률." + strKeyCount1
'            HONAQ001Data(i, 7) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.평가손익." + strKeyCount1
'            HONAQ001Data(i, 8) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.매입금액." + strKeyCount1
'            HONAQ001Data(i, 9) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.평가금액." + strKeyCount1
'            HONAQ001Data(i, 10) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.대주매각금." + strKeyCount1
'            HONAQ001Data(i, 11) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.대출일." + strKeyCount1
'            HONAQ001Data(i, 12) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.만기일." + strKeyCount1
'            HONAQ001Data(i, 13) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.현재가." + strKeyCount1
'            HONAQ001Data(i, 14) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.매수호가." + strKeyCount1
'            HONAQ001Data(i, 15) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.매도호가." + strKeyCount1
'            HONAQ001Data(i, 16) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.기준가." + strKeyCount1
'            HONAQ001Data(i, 17) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.종목코드." + strKeyCount1
'            HONAQ001Data(i, 18) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.등락부호." + strKeyCount1
'            HONAQ001Data(i, 19) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96007.신용만기일." + strKeyCount1
'            HONAQ001Data(i, 20) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'
'        Next i
'
''        For i = 0 To nCount1 - 1
''            Sheets("모듈").Cells(70 + i, 8).Value = HONAQ001Data(i, 0)
''            Sheets("모듈").Cells(70 + i, 9).Value = HONAQ001Data(i, 1)
''            Sheets("모듈").Cells(70 + i, 10).Value = HONAQ001Data(i, 2)
''            Sheets("모듈").Cells(70 + i, 11).Value = HONAQ001Data(i, 3)
''            Sheets("모듈").Cells(70 + i, 12).Value = HONAQ001Data(i, 4)
''            Sheets("모듈").Cells(70 + i, 13).Value = HONAQ001Data(i, 5)
''            Sheets("모듈").Cells(70 + i, 14).Value = HONAQ001Data(i, 6)
''            Sheets("모듈").Cells(70 + i, 15).Value = HONAQ001Data(i, 7)
''            Sheets("모듈").Cells(70 + i, 16).Value = HONAQ001Data(i, 8)
''            Sheets("모듈").Cells(70 + i, 17).Value = HONAQ001Data(i, 9)
''            Sheets("모듈").Cells(70 + i, 18).Value = HONAQ001Data(i, 10)
''            Sheets("모듈").Cells(70 + i, 19).Value = HONAQ001Data(i, 11)
''            Sheets("모듈").Cells(70 + i, 20).Value = HONAQ001Data(i, 12)
''            Sheets("모듈").Cells(70 + i, 21).Value = HONAQ001Data(i, 13)
''            Sheets("모듈").Cells(70 + i, 22).Value = HONAQ001Data(i, 14)
''            Sheets("모듈").Cells(70 + i, 23).Value = HONAQ001Data(i, 15)
''            Sheets("모듈").Cells(70 + i, 24).Value = HONAQ001Data(i, 16)
''            Sheets("모듈").Cells(70 + i, 25).Value = HONAQ001Data(i, 17)
''            Sheets("모듈").Cells(70 + i, 26).Value = HONAQ001Data(i, 18)
''            Sheets("모듈").Cells(70 + i, 27).Value = HONAQ001Data(i, 19)
''            Sheets("모듈").Cells(70 + i, 28).Value = HONAQ001Data(i, 20)
''
''        Next
'        For c = QuoteStockHStartPos To QuoteStockItemCount + QuoteStockHStartPos - 1
'            Sheets("가격수령").Cells(15, c).Value = 0
'            Sheets("가격수령").Cells(16, c).Value = 0
'        Next
'
'        For iRow = 0 To UBound(HONAQ001Data, 1) - 1
'            For c = QuoteStockHStartPos To QuoteStockItemCount + QuoteStockHStartPos - 1
'                If (Sheets("가격수령").Cells(3, c).Value = "A" + HONAQ001Data(iRow, 18)) Then
'                    Sheets("가격수령").Cells(15, c).Value = HONAQ001Data(iRow, 3)
'                    Sheets("가격수령").Cells(16, c).Value = CStr(CLng(HONAQ001Data(iRow, 9)) / CLng(HONAQ001Data(iRow, 3)))
'                End If
'            Next
'        Next
'        Application.Calculation = xlCalculationAutomatic
'
'    Case "96012": '주식:주문체결/전체 조회
'
'        strCount1 = Trim(WapiCommunityMngCtrl.GetTrData("96012.데이터카운트"))
'        If strCount1 <> "" Then
'            nCount1 = CInt(strCount1)
'        End If
'
'        If isSelectHONAQ003 Then
'            Application.Calculation = xlCalculationManual
'
'            For i = 0 To nCount1 - 1
'                strKeyCount1 = CStr(i)
'
'                strKeyName1 = "96012.주문번호." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 1).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.종목명." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 2).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.주문구분." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 3).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.주문수량." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 4).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.주문단가." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 5).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.체결량." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 6).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.체결단가." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 7).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.미체결수량." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 8).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.원주문번호." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 9).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.상태." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 10).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.주문시간." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 11).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.종목코드." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 12).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'                strKeyName1 = "96012.체결구분." + strKeyCount1
'                Sheets("체결내역").Cells(3 + i, 13).Value = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'
'            Next i
'
'            Application.Calculation = xlCalculationAutomatic
'        End If
'
'    Case "96013": '주식:주문 미체결 조회
'
'        strCount1 = Trim(WapiCommunityMngCtrl.GetTrData("96013.데이터카운트"))
'        If strCount1 <> "" Then
'            nCount1 = CInt(strCount1)
'        End If
'
'        ReDim HONAQ003Data(nCount1, 14) As String
'
'        If isSelectHONAQ003 Then
'            Application.Calculation = xlCalculationManual
'        End If
'
'        HONAQ003Count = nCount1
'
'        For i = 0 To nCount1 - 1
'            strKeyCount1 = CStr(i)
'            strKeyName1 = "96013.종목명." + strKeyCount1
'            HONAQ003Data(i, 0) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.주문구분ATTR." + strKeyCount1
'            HONAQ003Data(i, 1) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.주문구분." + strKeyCount1
'            HONAQ003Data(i, 2) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.주문수량." + strKeyCount1
'            HONAQ003Data(i, 3) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.미체결수량." + strKeyCount1
'            HONAQ003Data(i, 4) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.주문단가." + strKeyCount1
'            HONAQ003Data(i, 5) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.현재가." + strKeyCount1
'            HONAQ003Data(i, 6) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.주문번호." + strKeyCount1
'            HONAQ003Data(i, 7) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.원주문번호." + strKeyCount1
'            HONAQ003Data(i, 8) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.신용구분." + strKeyCount1
'            HONAQ003Data(i, 9) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.종목코드." + strKeyCount1
'            HONAQ003Data(i, 10) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.등락부호." + strKeyCount1
'            HONAQ003Data(i, 11) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.주문유형." + strKeyCount1
'            HONAQ003Data(i, 12) = Trim(WapiCommunityMngCtrl.GetTrData(strKeyName1))
'            strKeyName1 = "96013.주문조건." + strKeyCount1
'            HONAQ003Data(i, 13) = WapiCommunityMngCtrl.GetTrData(strKeyName1)
'
'
'            If isSelectHONAQ003 Then
'                Sheets("모듈").Cells(36 + i, 8).Value = HONAQ003Data(i, 0)
'                Sheets("모듈").Cells(36 + i, 9).Value = HONAQ003Data(i, 1)
'                Sheets("모듈").Cells(36 + i, 10).Value = HONAQ003Data(i, 2)
'                Sheets("모듈").Cells(36 + i, 11).Value = HONAQ003Data(i, 3)
'                Sheets("모듈").Cells(36 + i, 12).Value = HONAQ003Data(i, 4)
'                Sheets("모듈").Cells(36 + i, 13).Value = HONAQ003Data(i, 5)
'                Sheets("모듈").Cells(36 + i, 14).Value = HONAQ003Data(i, 6)
'                Sheets("모듈").Cells(36 + i, 15).Value = HONAQ003Data(i, 7)
'                Sheets("모듈").Cells(36 + i, 16).Value = HONAQ003Data(i, 8)
'                Sheets("모듈").Cells(36 + i, 17).Value = HONAQ003Data(i, 9)
'                Sheets("모듈").Cells(36 + i, 18).Value = HONAQ003Data(i, 10)
'                Sheets("모듈").Cells(36 + i, 19).Value = HONAQ003Data(i, 11)
'                Sheets("모듈").Cells(36 + i, 20).Value = HONAQ003Data(i, 12)
'                Sheets("모듈").Cells(36 + i, 21).Value = HONAQ003Data(i, 13)
'            End If
'
'        Next i
'        If isSelectHONAQ003 Then
'            Application.Calculation = xlCalculationAutomatic
'        End If
'
'    End Select
'
'    Sheets("모듈").Cells(9, 9).Value = "에러코드: " + WapiCommunityMngCtrl.GetErrCode() + ", 에러메세지: " + WapiCommunityMngCtrl.GetErrMsg()
End Sub

Private Sub WapiConnectMngCtrl_TrSocketStatus(ByVal nStatus As Long, ByVal nValue As Long)

End Sub




