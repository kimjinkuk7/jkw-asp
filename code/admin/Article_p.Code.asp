<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Article_p.Code.asp  -  게시글 관리 처리
'   mode : delete(논리삭제) | restore(복구) | notice(공지 토글) | purge(영구삭제)
'==============================================================================

Auth.CheckToken RequestUtil.ValueInit("_csrf", "")

Dim apMode, apId, apArr, apBack
apMode = LCase(RequestUtil.ValueInit("mode", ""))
apId   = RequestUtil.ValueIntInit("id", 0)

If apId <= 0 Then JsAlert "잘못된 접근입니다."

apArr = Article.AdminArticle_v(apId)
If Not IsArray(apArr) Then JsAlert "존재하지 않는 게시물입니다."

apBack = U("/admin/Article.asp") & "?bid=" & UE(RequestUtil.ValueInit("bid", "")) & _
         "&sw=" & UE(RequestUtil.ValueInit("sw", "")) & "&page=" & RequestUtil.ValueIntInit("page", 1)

Select Case apMode

    Case "delete"
        Article.Article_d apId
        JsAlertGo "삭제했습니다. (복구 가능)", apBack

    Case "restore"
        Article.Restore_u apId
        JsAlertGo "복구했습니다.", apBack

    Case "notice"
        Article.NoticeToggle_u apId
        JsGo apBack

    Case "purge"
        '--- 물리 파일까지 정리한 뒤 행 삭제 ---
        BoardFile.FileAll_d apId
        Article.Article_del_p apId
        JsAlertGo "영구 삭제했습니다.", apBack

    Case Else
        JsAlert "잘못된 접근입니다."

End Select
%>
