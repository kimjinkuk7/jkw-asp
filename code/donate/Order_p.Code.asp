<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<%
'==============================================================================
' code/donate/Order_p.Code.asp  -  주문 생성 (AJAX, 순수 텍스트)
'------------------------------------------------------------------------------
'   OK|<주문번호>|<금액>|<후원자명>|<PG>
'   FAIL|<메시지>
'
' 결제창(모달)을 열기 전에 주문을 먼저 만든다. 실제 PG 연동도 순서가 같다.
' 주문번호가 있어야 승인 단계에서 금액을 대조할 수 있다.
' ★ 금액은 여기서 검증한다. 화면 값을 그대로 믿으면 1원 결제가 가능해진다.
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
If Len(Auth.Token()) = 0 Then Out "FAIL|세션이 만료되었습니다. 새로고침 후 다시 시도해 주세요."
If CStr(Trim(Request.Form("_csrf") & "")) <> Auth.Token() Then Out "FAIL|잘못된 접근입니다."

Dim oAmount, oName, oMsg, oPublic, oOrderId, oNewId, oPg

oAmount = StringUtil.ToInt(Request.Form("amount"), 0)
oName   = Left(Trim(Request.Form("donor_name") & ""), 30)
oMsg    = Left(Trim(Request.Form("message")    & ""), 200)
oPublic = (Trim(Request.Form("is_public") & "") = "1")
oPg     = LCase(Left(Trim(Request.Form("pg") & ""), 20))

If Len(oName) = 0 Then Out "FAIL|후원자 이름을 입력해 주세요."

'--- 결제 모듈 검증. 허용 목록에 없는 값은 받지 않는다 ---
If Len(oPg) = 0 Then oPg = LCase(PAY_PROVIDER)
If InStr("," & LCase(Replace(PAY_PROVIDERS, " ", "")) & ",", "," & oPg & ",") = 0 Then
    Out "FAIL|사용할 수 없는 결제 모듈입니다."
End If

If oAmount < PAY_MIN_AMOUNT Then
    Out "FAIL|후원 금액은 " & StringUtil.Comma(PAY_MIN_AMOUNT) & "원 이상이어야 합니다."
End If
If oAmount > PAY_MAX_AMOUNT Then
    Out "FAIL|후원 금액은 " & StringUtil.Comma(PAY_MAX_AMOUNT) & "원 이하로 해주세요. 마음은 감사히 받겠습니다."
End If

oOrderId = Donate.NewOrderId()
oNewId   = Donate.Donate_i(oOrderId, GLB_MY_ID, oName, oMsg, oAmount, oPg, PAY_IS_TEST, oPublic)
If oNewId <= 0 Then Out "FAIL|주문 생성에 실패했습니다."

Donate.Log_i oOrderId, oPg, "request", 0, _
             "amount=" & oAmount & "&donor=" & oName, ""

Out "OK|" & oOrderId & "|" & oAmount & "|" & oName & "|" & oPg
%>
