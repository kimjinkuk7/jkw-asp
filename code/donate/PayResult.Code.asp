<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<%
'==============================================================================
' code/donate/PayResult.Code.asp  -  결제 결과 처리 + 결과 화면
'------------------------------------------------------------------------------
' mock : 이전 화면에서 고른 성공/실패를 그대로 적용한다.
' toss : 결제창이 돌려준 paymentKey 로 ★서버가 직접 승인 API를 호출★ 한다.
'        결제창이 성공을 알려줬다는 것만으로 승인된 것이 아니다.
'        승인 전에 금액을 반드시 우리 DB 값과 대조한다 (금액 조작 방어).
'==============================================================================
g_menuCode  = "donate"
g_bodyClass = "page-donate"
g_pageTitle = "결제 결과"

Dim prOrder, prArr, prStatus, prAmount, prName, prProvider
Dim prOk, prMsg, prMethod, prTid

prOrder = Left(Trim(RequestUtil.ValueInit("order", "")), 64)
If Len(prOrder) = 0 Then JsAlertGo "주문번호가 없습니다.", U("/donate/Default.asp")

prArr = Donate.Donate_v(prOrder)
If Not IsArray(prArr) Then JsAlertGo "주문을 찾을 수 없습니다.", U("/donate/Default.asp")

prStatus   = AV(prArr,  DN_STATUS, 0)
prAmount   = AVI(prArr, DN_AMOUNT, 0)
prName     = AV(prArr,  DN_NAME, 0)
prProvider = LCase(AV(prArr, DN_PROVIDER, 0))
prOk       = False
prMsg      = ""

'--- 이미 확정된 주문은 다시 처리하지 않는다 (새로고침/중복 콜백 방어) ---
If prStatus = "paid" Then
    prOk  = True
    prMsg = "이미 처리된 결제입니다."

ElseIf prStatus = "failed" Or prStatus = "canceled" Then
    prOk  = False
    prMsg = AV(prArr, DN_FAIL, 0)
    If Len(prMsg) = 0 Then prMsg = "결제가 완료되지 않았습니다."

'==============================================================================
' mock : 화면에서 고른 결과를 적용
'==============================================================================
ElseIf prProvider = "mock" Then

    If Not RequestUtil.IsPost() Then JsAlertGo "잘못된 접근입니다.", U("/donate/Default.asp")
    Auth.CheckToken Trim(Request.Form("_csrf") & "")

    If Trim(Request.Form("result") & "") = "success" Then
        prTid    = "MOCK-" & UCase(Left(StringUtil.NewKey(), 12))
        prMethod = "시연용"
        If Donate.Paid_u(prOrder, prTid, prMethod) Then
            prOk = True
        Else
            prOk  = False
            prMsg = "이미 처리된 주문입니다."
        End If
        Donate.Log_i prOrder, "mock", "result", 200, "result=success", "tid=" & prTid
    Else
        prMsg = "사용자가 결제를 취소했습니다. (시연용)"
        Donate.Failed_u prOrder, prMsg
        Donate.Log_i prOrder, "mock", "result", 200, "result=fail", prMsg
    End If

'==============================================================================
' toss : 서버에서 승인 API 호출
'==============================================================================
ElseIf prProvider = "toss" Then

    Dim tKey, tAmount, tUrl, tBody, tFailCode, tFailMsg

    ' 결제창이 실패로 돌아온 경우
    tFailCode = RequestUtil.ValueInit("code", "")
    tFailMsg  = RequestUtil.ValueInit("message", "")
    If Len(tFailCode) > 0 Then
        prMsg = tFailMsg
        If Len(prMsg) = 0 Then prMsg = "결제가 취소되었습니다. (" & tFailCode & ")"
        Donate.Failed_u prOrder, prMsg
        Donate.Log_i prOrder, "toss", "result", 0, "code=" & tFailCode, tFailMsg

    Else
        tKey    = Left(Trim(RequestUtil.ValueInit("paymentKey", "")), 100)
        tAmount = StringUtil.ToInt(RequestUtil.ValueInit("amount", "0"), 0)

        If Len(tKey) = 0 Then
            prMsg = "결제 키가 없습니다."
            Donate.Failed_u prOrder, prMsg

        ElseIf tAmount <> prAmount Then
            ' ★ 금액 대조. 결제창이 돌려준 금액이 우리가 만든 주문과 다르면 승인하지 않는다.
            prMsg = "결제 금액이 주문 금액과 다릅니다."
            Donate.Failed_u prOrder, prMsg
            Donate.Log_i prOrder, "toss", "error", 0, _
                         "expected=" & prAmount & " got=" & tAmount, ""

        ElseIf Len(TOSS_SECRET_KEY) = 0 Then
            prMsg = "서버에 토스 시크릿 키가 설정되지 않았습니다."
            Donate.Failed_u prOrder, prMsg

        Else
            tUrl  = "https://api.tosspayments.com/v1/payments/confirm"
            tBody = "{""paymentKey"":" & Json.Enc(tKey) & _
                    ",""orderId"":"    & Json.Enc(prOrder) & _
                    ",""amount"":"     & prAmount & "}"

            Http.Reset
            Http.TimeoutMs = 15000
            ' 토스는 시크릿 키를 아이디로, 비밀번호는 빈 값으로 하는 Basic 인증을 쓴다
            Http.SetBasicAuth TOSS_SECRET_KEY, ""

            If Not Http.PostJson(tUrl, tBody) Then
                prMsg = "승인 요청 실패 : " & Http.ErrMsg
                Donate.Failed_u prOrder, prMsg
                Donate.Log_i prOrder, "toss", "error", 0, tBody, Http.ErrMsg
            Else
                Donate.Log_i prOrder, "toss", "approve", Http.Status, tBody, Http.Body

                If Json.Parse(Http.Body) Then
                    If Http.IsOk And Json.Str("status", "") = "DONE" Then
                        prMethod = Json.Str("method", "카드")
                        If Donate.Paid_u(prOrder, tKey, prMethod) Then
                            prOk = True
                        Else
                            prOk  = True
                            prMsg = "이미 처리된 결제입니다."
                        End If
                    Else
                        prMsg = Json.Str("message", "승인이 거절되었습니다.")
                        Donate.Failed_u prOrder, prMsg
                    End If
                Else
                    prMsg = "승인 응답을 해석할 수 없습니다."
                    Donate.Failed_u prOrder, prMsg
                End If
            End If
        End If
    End If

Else
    prMsg = "지원하지 않는 결제 수단입니다."
    Donate.Failed_u prOrder, prMsg
End If

'--- 최신 상태 다시 읽기 (화면 표시용) ---
prArr = Donate.Donate_v(prOrder)
If IsArray(prArr) Then
    prStatus = AV(prArr,  DN_STATUS, 0)
    prMethod = AV(prArr,  DN_METHOD, 0)
    prTid    = AV(prArr,  DN_TID, 0)
    prOk     = (prStatus = "paid")
    If Not prOk And Len(prMsg) = 0 Then prMsg = AV(prArr, DN_FAIL, 0)
End If
%>
