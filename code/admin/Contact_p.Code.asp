<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Contact_p.Code.asp  -  문의 답변 저장 / 삭제
'==============================================================================

Auth.CheckToken RequestUtil.ValueInit("_csrf", "")

Dim xMode, xId
xMode = LCase(RequestUtil.ValueInit("mode", ""))
xId   = RequestUtil.ValueIntInit("id", 0)

If xId <= 0 Then JsAlert "잘못된 접근입니다."
If Not IsArray(Contact.Contact_v(xId)) Then JsAlert "존재하지 않는 문의입니다."

Select Case xMode

    Case "reply"
        Dim xReply
        xReply = Trim(Request.Form("reply_content") & "")
        If StringUtil.IsBlank(xReply) Then JsAlert "답변 내용을 입력해 주세요."
        Contact.Reply_u xId, xReply
        JsAlertGo "답변이 저장되었습니다.", U("/admin/Contact.asp") & "?id=" & xId

    Case "delete"
        Contact.Contact_d xId
        JsAlertGo "삭제되었습니다.", U("/admin/Contact.asp")

    Case Else
        JsAlert "잘못된 접근입니다."

End Select
%>
