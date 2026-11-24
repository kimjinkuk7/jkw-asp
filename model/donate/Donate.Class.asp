<%
'==============================================================================
' model/donate/Donate.Class.asp  -  개발자 지원금 (후원)
'------------------------------------------------------------------------------
' 상태 흐름
'   ready  주문 생성 (아직 결제 안 됨)
'     → paid      결제 승인 완료
'     → failed    승인 실패 / 사용자가 창을 닫음
'     → canceled  결제 후 취소
'
' 원칙
'   · 금액은 서버가 다시 계산한다. 화면에서 온 금액을 그대로 믿지 않는다.
'   · PG 통신 원문은 tb_payment_log 에 남긴다 (분쟁·디버깅에 필수).
'   · is_test 로 테스트 결제를 반드시 구분해 둔다.
'==============================================================================

'--- Donate_v() / Public_l() 결과 인덱스 ---
Const DN_ID       = 0
Const DN_ORDERID  = 1
Const DN_MEMBERID = 2
Const DN_NAME     = 3
Const DN_MESSAGE  = 4
Const DN_AMOUNT   = 5
Const DN_PROVIDER = 6
Const DN_METHOD   = 7
Const DN_STATUS   = 8
Const DN_TID      = 9
Const DN_FAIL     = 10
Const DN_ISTEST   = 11
Const DN_ISPUBLIC = 12
Const DN_PAIDDT   = 13
Const DN_REGDT    = 14

Class CDonate

    Public DBHelper

    Private Sub Class_Initialize()
        Set DBHelper = New CDBHelper
        DBHelper.ConnectionString = ConnectionString(DBCONN_JKWCOM)
    End Sub

    Private Function Cols()
        Cols = " donation_id, order_id, COALESCE(member_id,0), donor_name, COALESCE(message,''), " & _
               " amount, pg_provider, COALESCE(pay_method,''), pay_status, COALESCE(pg_tid,''), " & _
               " COALESCE(fail_reason,''), is_test, is_public, paid_dt, reg_dt "
    End Function

    '==========================================================================
    ' 주문번호
    '==========================================================================
    '<Desc>주문번호. 날짜 + 난수. PG 들이 요구하는 6~64자 영숫자 범위를 지킨다.</Desc>
    Public Function NewOrderId()
        Dim d
        d = Now()
        ' NewKey() 는 "날짜14자 + 난수18자" 형태다. 앞을 자르면 난수가 안 들어가므로
        ' 뒤에서 잘라 쓴다. (같은 초에 두 건이 들어오면 order_id 유니크 제약에 걸린다)
        NewOrderId = "JKW" & Year(d) & P2(Month(d)) & P2(Day(d)) & _
                     P2(Hour(d)) & P2(Minute(d)) & P2(Second(d)) & _
                     UCase(Right(StringUtil.NewKey(), 10))
    End Function

    Private Function P2(ByVal n)
        P2 = Right("0" & CStr(n), 2)
    End Function

    '==========================================================================
    ' 등록 / 조회
    '==========================================================================
    Public Function Donate_i(ByVal pOrderId, ByVal pMemberId, ByVal pName, ByVal pMessage, _
                             ByVal pAmount, ByVal pProvider, ByVal pIsTest, ByVal pIsPublic)
        With DBHelper
            .Sql = "INSERT INTO tb_donation " & _
                   " (order_id, member_id, donor_name, message, amount, pg_provider, " & _
                   "  pay_status, is_test, is_public, reg_ip) " & _
                   "VALUES (?, ?, ?, ?, ?, ?, 'ready', ?, ?, ?)"
            .SetParamVarchar(pOrderId)
            .SetParamInt(NullIfZero(pMemberId))
            .SetParamVarchar(pName)
            .SetParamVarchar(pMessage)
            .SetParamInt(pAmount)
            .SetParamVarchar(LCase(CStr(pProvider)))
            .SetParamBit(pIsTest)
            .SetParamBit(pIsPublic)
            .SetParamVarchar(RequestUtil.ClientIp())
        End With
        Donate_i = DBHelper.ExecuteIdentity()
    End Function

    Public Function Donate_v(ByVal pOrderId)
        With DBHelper
            .Sql = "SELECT " & Cols() & " FROM tb_donation WHERE order_id = ?"
            .SetParamVarchar(pOrderId)
        End With
        Donate_v = DBHelper.ExecuteGetRows()
    End Function

    '==========================================================================
    ' 상태 변경
    '==========================================================================
    '<Desc>승인 완료. ready 상태만 paid 로 바꾼다 (중복 승인 방지).</Desc>
    Public Function Paid_u(ByVal pOrderId, ByVal pTid, ByVal pMethod)
        With DBHelper
            .Sql = "UPDATE tb_donation " & _
                   "   SET pay_status = 'paid', pg_tid = ?, pay_method = ?, " & _
                   "       paid_dt = CURRENT_TIMESTAMP " & _
                   " WHERE order_id = ? AND pay_status = 'ready'"
            .SetParamVarchar(pTid)
            .SetParamVarchar(pMethod)
            .SetParamVarchar(pOrderId)
        End With
        DBHelper.Execute()
        Paid_u = (DBHelper.Affected > 0)
    End Function

    Public Sub Failed_u(ByVal pOrderId, ByVal pReason)
        With DBHelper
            .Sql = "UPDATE tb_donation SET pay_status = 'failed', fail_reason = ? " & _
                   " WHERE order_id = ? AND pay_status = 'ready'"
            .SetParamVarchar(Left(CStr(pReason & ""), 200))
            .SetParamVarchar(pOrderId)
        End With
        DBHelper.Execute()
    End Sub

    '==========================================================================
    ' 공개 목록 / 집계
    '==========================================================================
    '<Desc>후원자 목록 (승인 완료 + 공개 동의 건만)</Desc>
    Public Function Public_l(ByVal pTop)
        With DBHelper
            .Sql = "SELECT TOP (?) " & Cols() & _
                   "  FROM tb_donation " & _
                   " WHERE pay_status = 'paid' AND is_public = 1 " & _
                   " ORDER BY donation_id DESC"
            .SetParamInt(pTop)
        End With
        Public_l = DBHelper.ExecuteGetRows()
    End Function

    '<Desc>집계  arr(0)=건수 arr(1)=합계</Desc>
    Public Function Stat_v()
        With DBHelper
            .Sql = "SELECT COUNT(*), COALESCE(SUM(amount),0) " & _
                   "  FROM tb_donation WHERE pay_status = 'paid'"
        End With
        Stat_v = DBHelper.ExecuteGetRows()
    End Function

    '<Desc>관리자용 전체 목록</Desc>
    Public Function Admin_l(ByVal pOffset, ByVal pRows)
        With DBHelper
            .Sql = "SELECT " & Cols() & " FROM tb_donation " & _
                   " ORDER BY donation_id DESC OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
            .SetParamInt(pOffset)
            .SetParamInt(pRows)
        End With
        Admin_l = DBHelper.ExecuteGetRows()
    End Function

    '==========================================================================
    ' PG 통신 로그
    '==========================================================================
    Public Sub Log_i(ByVal pOrderId, ByVal pProvider, ByVal pStep, ByVal pHttpStatus, _
                     ByVal pReq, ByVal pRes)
        On Error Resume Next                   ' 로그 실패가 결제를 막아선 안 된다
        With DBHelper
            .Sql = "INSERT INTO tb_payment_log " & _
                   " (order_id, pg_provider, step, http_status, req_body, res_body) " & _
                   "VALUES (?, ?, ?, ?, ?, ?)"
            .SetParamVarchar(pOrderId)
            .SetParamVarchar(LCase(CStr(pProvider)))
            .SetParamVarchar(pStep)
            .SetParamInt(pHttpStatus)
            .SetParamText(Left(CStr(pReq & ""), 4000))
            .SetParamText(Left(CStr(pRes & ""), 4000))
        End With
        DBHelper.Execute()
        On Error GoTo 0
    End Sub

    '==========================================================================
    ' 내부
    '==========================================================================
    Private Function NullIfZero(ByVal v)
        If IsNull(v) Or IsEmpty(v) Then NullIfZero = Null : Exit Function
        If Not StringUtil.IsNum(v) Then NullIfZero = Null : Exit Function
        If CLng(CDbl(CStr(v))) <= 0 Then NullIfZero = Null Else NullIfZero = CLng(CDbl(CStr(v)))
    End Function

End Class

Dim Donate
Set Donate = New CDonate
%>
