<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<%
'==============================================================================
' code/donate/Pay.Code.asp  -  결제 진행 화면
'------------------------------------------------------------------------------
' mock : 결제창 대신 성공/실패를 고르는 시연용 화면을 보여준다.
' toss : 토스페이먼츠 결제창(JS SDK)을 띄운다. 성공하면 PayResult 로 돌아온다.
'==============================================================================
g_menuCode  = "donate"
g_bodyClass = "page-donate"
g_pageTitle = "결제"

Dim pyOrderId, pyArr, pyAmount, pyName, pyStatus, pyProvider

pyOrderId = Left(Trim(RequestUtil.ValueInit("order", "")), 64)
If Len(pyOrderId) = 0 Then JsAlertGo "주문번호가 없습니다.", U("/donate/Default.asp")

pyArr = Donate.Donate_v(pyOrderId)
If Not IsArray(pyArr) Then JsAlertGo "주문을 찾을 수 없습니다.", U("/donate/Default.asp")

pyAmount   = AVI(pyArr, DN_AMOUNT, 0)
pyName     = AV(pyArr,  DN_NAME, 0)
pyStatus   = AV(pyArr,  DN_STATUS, 0)
pyProvider = AV(pyArr,  DN_PROVIDER, 0)

'--- 이미 처리된 주문이면 결과 화면으로 ---
If pyStatus <> "ready" Then
    Response.Redirect U("/donate/PayResult.asp") & "?order=" & UE(pyOrderId)
End If

'--- 토스 사용 시 키 확인 ---
Dim pyTossReady
pyTossReady = (LCase(pyProvider) = "toss" And Len(TOSS_CLIENT_KEY) > 0)

If LCase(pyProvider) = "toss" And Not pyTossReady Then
    JsAlertGo "토스페이먼츠 키가 설정되지 않았습니다. App.Conf.asp 를 확인해 주세요.", U("/donate/Default.asp")
End If
%>
