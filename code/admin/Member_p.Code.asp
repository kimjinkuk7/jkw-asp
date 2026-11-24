<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Member_p.Code.asp  -  회원 관리 처리
'   mode : level(등급/사용여부) | unlock(로그인 잠금 해제)
'==============================================================================

Auth.CheckToken RequestUtil.ValueInit("_csrf", "")

Dim mpMode, mpId, mpArr
mpMode = LCase(RequestUtil.ValueInit("mode", ""))
mpId   = RequestUtil.ValueIntInit("id", 0)

If mpId <= 0 Then JsAlert "잘못된 접근입니다."
If mpId = GLB_MY_ID Then JsAlert "본인 계정은 이 화면에서 변경할 수 없습니다."

mpArr = Member.Member_v(mpId)
If Not IsArray(mpArr) Then JsAlert "존재하지 않는 회원입니다."

Dim mpBack
mpBack = U("/admin/Member.asp") & "?sw=" & UE(RequestUtil.ValueInit("sw", "")) & _
         "&lv=" & RequestUtil.ValueIntInit("lv", -1) & "&page=" & RequestUtil.ValueIntInit("page", 1)

Select Case mpMode

    Case "level"
        Dim newLv, newUse
        newLv  = StringUtil.ToInt(Request.Form("member_level"), 3)
        newUse = (Trim(Request.Form("is_use") & "") = "1")

        If InStr(",0,1,3,5,9,", "," & newLv & ",") = 0 Then JsAlert "허용되지 않는 등급입니다."

        Member.Level_u mpId, newLv, newUse
        JsGo mpBack

    Case "unlock"
        Member.Unlock_u mpId
        JsAlertGo "로그인 잠금을 해제했습니다.", mpBack

    Case Else
        JsAlert "잘못된 접근입니다."

End Select
%>
