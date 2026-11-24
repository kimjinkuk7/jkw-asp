<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<!--#include file="../../model/donate/PgInicis.Class.asp"-->
<%
'==============================================================================
' code/donate/InicisReturn.Code.asp  -  이니시스 결제창 리턴 처리
'------------------------------------------------------------------------------
' 이니시스가 결제창에서 POST 로 돌려보낸다. 여기서 ★서버가 승인 요청★ 을 한다.
' 결제창이 성공을 알려준 것만으로 승인이 아니다.
'
' 받는 값 : resultCode, resultMsg, authToken, authUrl, netCancelUrl, orderNumber, ...
'==============================================================================

Dim irCode, irMsg, irToken, irAuthUrl, irOrder
Dim irArr, irAmount, irStatus, irTid, irMethod, irErr

irCode    = Trim(Request.Form("resultCode")  & "")
irMsg     = Left(Trim(Request.Form("resultMsg") & ""), 200)
irToken   = Trim(Request.Form("authToken")   & "")
irAuthUrl = Trim(Request.Form("authUrl")     & "")
irOrder   = Left(Trim(Request.Form("orderNumber") & ""), 64)

If Len(irOrder) = 0 Then irOrder = Left(Trim(Request.Form("oid") & ""), 64)
If Len(irOrder) = 0 Then JsAlertGo "주문번호를 확인할 수 없습니다.", U("/donate/Default.asp")

irArr = Donate.Donate_v(irOrder)
If Not IsArray(irArr) Then JsAlertGo "주문을 찾을 수 없습니다.", U("/donate/Default.asp")

irAmount = AVI(irArr, DN_AMOUNT, 0)
irStatus = AV(irArr,  DN_STATUS, 0)

Donate.Log_i irOrder, "inicis", "result", 0, _
             "resultCode=" & irCode & "&hasToken=" & CStr(Len(irToken) > 0), irMsg

'--- 이미 끝난 주문이면 결과 화면으로 ---
If irStatus <> "ready" Then
    Response.Redirect U("/donate/PayResult.asp") & "?order=" & UE(irOrder)
End If

'--- 결제창에서 실패/취소한 경우 ---
If irCode <> "0000" Then
    If Len(irMsg) = 0 Then irMsg = "결제가 취소되었습니다. (" & irCode & ")"
    Donate.Failed_u irOrder, irMsg
    Response.Redirect U("/donate/PayResult.asp") & "?order=" & UE(irOrder)
End If

If Len(irToken) = 0 Or Len(irAuthUrl) = 0 Then
    Donate.Failed_u irOrder, "승인 정보가 없습니다."
    Response.Redirect U("/donate/PayResult.asp") & "?order=" & UE(irOrder)
End If

'--- ★ 서버 승인 ---
If PgInicis.Approve(irAuthUrl, irToken, irOrder, irAmount, irTid, irMethod, irErr) Then
    If Not Donate.Paid_u(irOrder, irTid, irMethod) Then
        ' 이미 paid 로 바뀐 경우(중복 콜백). 실패로 보지 않는다.
    End If
Else
    Donate.Failed_u irOrder, irErr
End If

Response.Redirect U("/donate/PayResult.asp") & "?order=" & UE(irOrder)
%>
