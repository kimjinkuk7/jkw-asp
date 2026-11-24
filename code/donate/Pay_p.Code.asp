<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<%
'==============================================================================
' code/donate/Pay_p.Code.asp  -  후원 주문 생성
'------------------------------------------------------------------------------
' 화면에서 온 값을 검증하고 tb_donation 에 ready 상태로 넣은 뒤 결제 단계로 보낸다.
' ★ 금액은 여기서 다시 검증한다. 화면 값을 그대로 믿으면 1원 결제가 가능해진다.
'==============================================================================

If Not RequestUtil.IsPost() Then JsAlert "잘못된 접근입니다."
Auth.CheckToken Trim(Request.Form("_csrf") & "")

Dim pAmount, pName, pMsg, pPublic, pOrderId, pNewId, pPg

pAmount = StringUtil.ToInt(Request.Form("amount"), 0)
pName   = Left(Trim(Request.Form("donor_name") & ""), 30)
pMsg    = Left(Trim(Request.Form("message")    & ""), 200)
pPublic = (Trim(Request.Form("is_public") & "") = "1")
pPg     = LCase(Left(Trim(Request.Form("pg") & ""), 20))

If Len(pName) = 0 Then JsAlert "후원자 이름을 입력해 주세요."

If Len(pPg) = 0 Then pPg = LCase(PAY_PROVIDER)
If InStr("," & LCase(Replace(PAY_PROVIDERS, " ", "")) & ",", "," & pPg & ",") = 0 Then
    JsAlert "사용할 수 없는 결제 모듈입니다."
End If

If pAmount < PAY_MIN_AMOUNT Then
    JsAlert "후원 금액은 " & StringUtil.Comma(PAY_MIN_AMOUNT) & "원 이상이어야 합니다."
End If
If pAmount > PAY_MAX_AMOUNT Then
    JsAlert "후원 금액은 " & StringUtil.Comma(PAY_MAX_AMOUNT) & "원 이하로 해주세요." & vbCrLf & _
            "마음은 감사히 받겠습니다."
End If

'--- 주문 생성 ---
pOrderId = Donate.NewOrderId()
pNewId   = Donate.Donate_i(pOrderId, GLB_MY_ID, pName, pMsg, pAmount, pPg, PAY_IS_TEST, pPublic)
If pNewId <= 0 Then JsAlert "주문 생성에 실패했습니다. 잠시 후 다시 시도해 주세요."

Donate.Log_i pOrderId, pPg, "request", 0, _
             "amount=" & pAmount & "&donor=" & pName, ""

'--- 결제 단계로 (모듈별 진입점) ---
If pPg = "inicis" Then
    Response.Redirect U("/donate/InicisPay.asp") & "?order=" & UE(pOrderId)
Else
    Response.Redirect U("/donate/Pay.asp") & "?order=" & UE(pOrderId)
End If
%>
