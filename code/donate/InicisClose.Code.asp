<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<%
'==============================================================================
' code/donate/InicisClose.Code.asp  -  이니시스 결제창 닫기(closeUrl)
'------------------------------------------------------------------------------
' 이니시스는 결제창을 닫을 때 closeUrl 을 iframe 안에서 호출한다.
' 여기에 일반 페이지를 지정하면 사이트 전체가 결제창 iframe 안에 그려져
' 이상한 화면이 된다. 그래서 이 페이지는 레이아웃을 쓰지 않고
' "부모(최상위) 창을 옮기는 스크립트" 만 내보낸다.
'
' 닫기는 사용자 취소이므로 ready 상태면 failed 로 정리한다.
' (이미 승인된 주문은 Failed_u 가 건드리지 않는다)
'==============================================================================

Dim icOrder, icArr, icGo

icOrder = Left(Trim(RequestUtil.ValueInit("order", "")), 64)
icGo    = U("/donate/Default.asp")

If Len(icOrder) > 0 Then
    icArr = Donate.Donate_v(icOrder)
    If IsArray(icArr) Then
        If AV(icArr, DN_STATUS, 0) = "ready" Then
            Donate.Failed_u icOrder, "사용자가 결제창을 닫았습니다."
            Donate.Log_i icOrder, "inicis", "result", 0, "closeUrl", "사용자 취소"
        End If
        icGo = U("/donate/PayResult.asp") & "?order=" & UE(icOrder)
    End If
End If

Response.Clear
Response.ContentType = "text/html"
Response.CharSet     = "utf-8"
Response.Expires     = -1
Response.Write "<!DOCTYPE html><html lang=""ko""><head><meta charset=""utf-8"">" & _
               "<title>결제 취소</title></head>" & _
               "<body style=""font:14px sans-serif;padding:20px"">결제창을 닫았습니다…" & _
               "<script>(function(){" & _
               "  var u='" & JE(icGo) & "';" & _
               "  try{ if(window.top && window.top!==window){ window.top.location.replace(u); return; } }catch(e){}" & _
               "  try{ if(parent && parent!==window){ parent.location.replace(u); return; } }catch(e){}" & _
               "  location.replace(u);" & _
               "})();</script></body></html>"
Response.End
%>
