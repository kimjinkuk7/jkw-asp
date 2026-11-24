<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<%
'==============================================================================
' code/donate/Cancel_p.Code.asp  -  결제창을 닫았을 때 주문 정리 (AJAX)
'------------------------------------------------------------------------------
'   OK|CANCELED   또는   OK|SKIP (이미 끝난 주문)
'   FAIL|<메시지>
'
' 사용자가 PG 결제창을 그냥 닫으면 주문이 ready 로 남는다. 그대로 두면
' "결제하려다 만" 주문이 계속 쌓여 무슨 일이 있었는지 알 수 없다.
' ready 인 것만 failed 로 바꾸므로, 이미 승인된 주문은 건드리지 않는다.
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

Dim cOrder, cArr, cReason

cOrder = Left(Trim(Request.Form("order") & ""), 64)
If Len(cOrder) = 0 Then Out "FAIL|주문번호가 없습니다."

cArr = Donate.Donate_v(cOrder)
If Not IsArray(cArr) Then Out "FAIL|주문을 찾을 수 없습니다."

If AV(cArr, DN_STATUS, 0) <> "ready" Then Out "OK|SKIP"

cReason = Left(Trim(Request.Form("reason") & ""), 200)
If Len(cReason) = 0 Then cReason = "사용자가 결제창을 닫았습니다."

Donate.Failed_u cOrder, cReason
Donate.Log_i cOrder, AV(cArr, DN_PROVIDER, 0), "result", 0, "canceled", cReason

Out "OK|CANCELED"
%>
