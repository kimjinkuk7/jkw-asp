<!--#include file="Board.Code.asp"-->
<%
'==============================================================================
' code/board/View.Code.asp  -  게시글 상세 컨트롤러
'==============================================================================

'--- 열람 권한 ---
If Not bdCanRead Then Auth.CheckLevel AVI(bdArr, BM_AUTH_READ, 0)

'--- 글 조회 ---
Dim artId, artArr
artId = RequestUtil.ValueIntInit("id", 0)
If artId <= 0 Then Go404 "잘못된 접근입니다."

artArr = Article.Article_v(artId)
If Not IsArray(artArr) Then Go404 "삭제되었거나 존재하지 않는 게시물입니다."

'--- 게시판 불일치 방지 (다른 게시판 글을 이 게시판으로 열람 시도) ---
If AV(artArr, AVW_BID, 0) <> bdId Then Go404 "잘못된 접근입니다."

Dim artTitle, artContent, artWriterId, artWriterName, artRegDt, artUpdDt
Dim artHit, artCmtCnt, artFileCnt, artTags, artLink, artCateName
Dim artIsNotice, artIsSecret, artHasPw

artTitle      = AV(artArr,  AVW_TITLE, 0)
artContent    = AV(artArr,  AVW_CONTENT, 0)
artWriterId   = AVI(artArr, AVW_WID, 0)
artWriterName = AV(artArr,  AVW_WNAME, 0)
artRegDt      = AV(artArr,  AVW_REGDT, 0)
artUpdDt      = AV(artArr,  AVW_UPDDT, 0)
artHit        = AVI(artArr, AVW_HIT, 0)
artCmtCnt     = AVI(artArr, AVW_CMT, 0)
artFileCnt    = AVI(artArr, AVW_FILE, 0)
artTags       = AV(artArr,  AVW_TAGS, 0)
artLink       = AV(artArr,  AVW_LINK, 0)
artCateName   = AV(artArr,  AVW_CNAME, 0)
artIsNotice   = AVB(artArr, AVW_NOTICE, 0)
artIsSecret   = AVB(artArr, AVW_SECRET, 0)
artHasPw      = AVB(artArr, AVW_HASPW, 0)

'--- 비밀글 접근 통제 ---------------------------------------------------------
'    관리자 / 작성자 본인 / 비밀번호 확인을 통과한 세션만 열람 가능
Dim pwOkKey, isPwOk
pwOkKey = "PWOK_" & bdId & "_" & artId
isPwOk  = (Session(pwOkKey) & "" = "1")

If artIsSecret Then
    If Not (bdCanAdmin Or (Auth.MyId() > 0 And Auth.MyId() = artWriterId) Or isPwOk) Then
        Response.Redirect U("/board/PwCheck.asp") & "?bid=" & UE(bdId) & "&id=" & artId & "&act=read"
    End If
End If

'--- 조회수 증가 (같은 세션 중복 제외) ----------------------------------------
Dim hitKey, hitLog
hitKey = "|" & bdId & ":" & artId & "|"
hitLog = Session("HITLOG") & ""
If InStr(hitLog, hitKey) = 0 Then
    Article.HitUp artId
    artHit = artHit + 1
    If Len(hitLog) > 3000 Then hitLog = ""
    Session("HITLOG") = hitLog & hitKey
End If

'--- 첨부파일 ---
Dim fileArr, fileCnt
fileCnt = 0
If bdUseFile Then
    fileArr = BoardFile.File_l(artId)
    fileCnt = RowCnt(fileArr)
End If

'--- 댓글 ---
Dim cmtArr, cmtCnt
cmtCnt = 0
If bdUseComment Then
    cmtArr = Comment.Comment_l(artId)
    cmtCnt = RowCnt(cmtArr)
End If

'--- 이전 / 다음 글 ---
Dim prevArr, nextArr
prevArr = Article.Sibling_v(bdId, artId, "prev")
nextArr = Article.Sibling_v(bdId, artId, "next")

'--- 버튼 노출 판정 ---
Dim canEdit, canReply
canEdit  = CanEditArticle(artWriterId, artHasPw)
canReply = (bdUseReply And bdCanWrite)

'--- 링크 ---
Dim backUrl, modifyUrl, deleteUrl, replyUrl
backUrl   = BoardBackUrl()
modifyUrl = U("/board/Write.asp")  & "?bid=" & UE(bdId) & "&id=" & artId & "&mode=modify" & ListStateQs()
replyUrl  = U("/board/Write.asp")  & "?bid=" & UE(bdId) & "&id=" & artId & "&mode=reply"  & ListStateQs()
deleteUrl = U("/board/Delete_p.asp") & "?bid=" & UE(bdId) & "&id=" & artId & "&_csrf=" & UE(Auth.Token())

'--- 메타 ---
g_pageTitle = artTitle
g_pageDesc  = StringUtil.Cut(StringUtil.StripTags(artContent), 150)
%>
