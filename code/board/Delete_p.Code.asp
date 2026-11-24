<!--#include file="Board.Code.asp"-->
<%
'==============================================================================
' code/board/Delete_p.Code.asp  -  게시글 삭제 처리 (논리 삭제)
'------------------------------------------------------------------------------
' 회원 글   : 본인 또는 게시판 관리자
' 비회원 글 : PwCheck.asp 를 통과한 세션만
'==============================================================================

Auth.CheckToken RequestUtil.ValueInit("_csrf", "")

Dim artId, artArr, aWid, aHasPw, okDel
artId = RequestUtil.ValueIntInit("id", 0)
If artId <= 0 Then JsAlert "잘못된 접근입니다."

artArr = Article.Article_v(artId)
If Not IsArray(artArr) Then JsAlert "이미 삭제되었거나 존재하지 않는 게시물입니다."
If AV(artArr, AVW_BID, 0) <> bdId Then JsAlert "잘못된 접근입니다."

aWid   = AVI(artArr, AVW_WID, 0)
aHasPw = AVB(artArr, AVW_HASPW, 0)

okDel = False
If bdCanAdmin Then okDel = True
If Auth.MyId() > 0 And Auth.MyId() = aWid Then okDel = True
If aWid = 0 And aHasPw And (Session("PWOK_" & bdId & "_" & artId) & "" = "1") Then okDel = True

If Not okDel Then JsAlert "삭제 권한이 없습니다."

'--- 논리 삭제 (관리자 화면에서 복구 가능) ---
Article.Article_d artId

'--- 인증 세션 정리 ---
Session("PWOK_" & bdId & "_" & artId) = ""

JsAlertGo "삭제되었습니다.", U("/board/List.asp") & "?bid=" & UE(bdId) & "&page=" & RequestUtil.ValueIntInit("page", 1)
%>
