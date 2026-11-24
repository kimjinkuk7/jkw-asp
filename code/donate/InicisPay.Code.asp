<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<!--#include file="../../model/donate/PgInicis.Class.asp"-->
<%
'==============================================================================
' code/donate/InicisPay.Code.asp  -  이니시스 결제창 호출 준비
'------------------------------------------------------------------------------
' 서명이 서버에서만 만들어져야 하므로, 주문을 만든 뒤 이 페이지로 와서
' hidden form 을 렌더링하고 INIStdPay.pay() 를 호출한다.
' 결제창 자체는 이니시스가 iframe 모달로 띄운다.
'==============================================================================
g_menuCode  = "donate"
g_bodyClass = "page-donate"
g_pageTitle = "결제"

Dim ipOrder, ipArr, ipAmount, ipName, ipStatus
Dim ipTs, ipSign, ipVerify, ipMKey, ipReturnUrl, ipCloseUrl

ipOrder = Left(Trim(RequestUtil.ValueInit("order", "")), 64)
If Len(ipOrder) = 0 Then JsAlertGo "주문번호가 없습니다.", U("/donate/Default.asp")

ipArr = Donate.Donate_v(ipOrder)
If Not IsArray(ipArr) Then JsAlertGo "주문을 찾을 수 없습니다.", U("/donate/Default.asp")

ipStatus = AV(ipArr,  DN_STATUS, 0)
ipAmount = AVI(ipArr, DN_AMOUNT, 0)
ipName   = AV(ipArr,  DN_NAME, 0)

If LCase(AV(ipArr, DN_PROVIDER, 0)) <> "inicis" Then
    JsAlertGo "결제 수단이 맞지 않습니다.", U("/donate/Default.asp")
End If

'--- 이미 끝난 주문이면 결과로 ---
If ipStatus <> "ready" Then
    Response.Redirect U("/donate/PayResult.asp") & "?order=" & UE(ipOrder)
End If

'--- 서명 생성 (모두 서버에서) ---
ipTs        = PgInicis.Timestamp()
ipSign      = PgInicis.Signature(ipOrder, ipAmount, ipTs)
ipVerify    = PgInicis.Verification(ipOrder, ipAmount, ipTs)
ipMKey      = PgInicis.MKey()
ipReturnUrl = GLB_SITE_URL & U("/donate/InicisReturn.asp")
' closeUrl 은 결제창 iframe 안에서 열린다. 일반 페이지를 주면 사이트 전체가
' 결제창 안에 그려져 이상한 화면이 되므로, 부모 창만 옮기는 전용 페이지를 쓴다.
ipCloseUrl  = GLB_SITE_URL & U("/donate/InicisClose.asp") & "?order=" & UE(ipOrder)

Donate.Log_i ipOrder, "inicis", "request", 0, _
             "oid=" & ipOrder & "&price=" & ipAmount & "&timestamp=" & ipTs, ""
%>
