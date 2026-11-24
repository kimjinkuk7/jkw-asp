<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<%
'==============================================================================
' code/donate/Approve_p.Code.asp  -  결제 승인 (AJAX, 순수 텍스트)
'------------------------------------------------------------------------------
'   OK|<거래번호>|<결제수단>
'   FAIL|<메시지>
'
' 모의 결제창에서 "결제하기" 를 누르면 여기로 온다.
' 실제 PG 라면 이 자리에서 PG 승인 API 를 호출한다 (toss 는 PayResult 가 담당).
'
' 승인은 항상 서버가 한다. 결제창이 성공을 알려준 것만으로 승인이 아니다.
' ready 상태만 paid 로 바꾸므로 중복 요청이 들어와도 두 번 승인되지 않는다.
'==============================================================================

Response.Clear
Response.ContentType = "text/plain"
Response.CharSet     = "utf-8"
Response.Expires     = -1

Sub Out(ByVal s)
    Response.Write s
    Response.End
End Sub

If Not RequestUtil.IsPost() Then Out "FAIL|잘못된 접근입니다."
If Len(Auth.Token()) = 0 Then Out "FAIL|세션이 만료되었습니다."
If CStr(Trim(Request.Form("_csrf") & "")) <> Auth.Token() Then Out "FAIL|잘못된 접근입니다."

Dim apOrder, apResult, apArr, apStatus, apProvider, apTid, apMethod

apOrder  = Left(Trim(Request.Form("order")  & ""), 64)
apResult = LCase(Trim(Request.Form("result") & ""))

If Len(apOrder) = 0 Then Out "FAIL|주문번호가 없습니다."

apArr = Donate.Donate_v(apOrder)
If Not IsArray(apArr) Then Out "FAIL|주문을 찾을 수 없습니다."

apStatus   = AV(apArr, DN_STATUS, 0)
apProvider = LCase(AV(apArr, DN_PROVIDER, 0))

If apStatus = "paid" Then Out "OK|" & AV(apArr, DN_TID, 0) & "|" & AV(apArr, DN_METHOD, 0)
If apStatus <> "ready" Then Out "FAIL|이미 종료된 주문입니다."

' 이 엔드포인트는 모의 결제창 전용이다.
' 실제 PG 는 결제창이 돌려준 키로 PayResult 에서 승인한다.
If apProvider <> "mock" Then Out "FAIL|이 결제 수단은 이 경로로 승인하지 않습니다."

If apResult = "fail" Then
    Donate.Failed_u apOrder, "사용자가 결제를 취소했습니다. (시연용)"
    Donate.Log_i apOrder, "mock", "result", 200, "result=fail", ""
    Out "FAIL|결제가 취소되었습니다."
End If

apTid    = "MOCK-" & UCase(Right(StringUtil.NewKey(), 12))
apMethod = Left(Trim(Request.Form("method") & ""), 30)
If Len(apMethod) = 0 Then apMethod = "카드(시연용)"

If Not Donate.Paid_u(apOrder, apTid, apMethod) Then
    Out "FAIL|이미 처리된 주문입니다."
End If

Donate.Log_i apOrder, "mock", "approve", 200, "result=success&method=" & apMethod, "tid=" & apTid

Out "OK|" & apTid & "|" & apMethod
%>
