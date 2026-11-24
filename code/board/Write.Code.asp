<!--#include file="Board.Code.asp"-->
<%
'==============================================================================
' code/board/Write.Code.asp  -  글쓰기 / 수정 / 답변 폼 컨트롤러
'------------------------------------------------------------------------------
' mode : write(신규) | modify(수정) | reply(답변)
'==============================================================================

Dim wMode, artId, artArr
wMode = LCase(RequestUtil.ValueInit("mode", "write"))
If wMode <> "modify" And wMode <> "reply" Then wMode = "write"

artId = RequestUtil.ValueIntInit("id", 0)

'--- 쓰기 권한 ---
If Not bdCanWrite Then Auth.CheckLevel AVI(bdArr, BM_AUTH_WRITE, 0)

'--- 답변 기능 확인 ---
If wMode = "reply" And Not bdUseReply Then JsAlert "이 게시판은 답변 기능을 사용하지 않습니다."

'--- 폼 초기값 ---
Dim fTitle, fContent, fCateId, fTags, fLinkUrl, fWriterName, fIsNotice, fIsSecret
Dim artWriterId, artHasPw
fTitle      = ""
fContent    = ""
fCateId     = RequestUtil.ValueIntInit("cid", 0)
fTags       = ""
fLinkUrl    = ""
fWriterName = ""
fIsNotice   = False
fIsSecret   = False
artWriterId = 0
artHasPw    = False

'--- 기존 글 로드 (수정 / 답변) ---
If wMode = "modify" Or wMode = "reply" Then

    If artId <= 0 Then JsAlert "잘못된 접근입니다."

    artArr = Article.Article_v(artId)
    If Not IsArray(artArr) Then JsAlert "삭제되었거나 존재하지 않는 게시물입니다."
    If AV(artArr, AVW_BID, 0) <> bdId Then JsAlert "잘못된 접근입니다."

    artWriterId = AVI(artArr, AVW_WID, 0)
    artHasPw    = AVB(artArr, AVW_HASPW, 0)

    If wMode = "modify" Then
        '--- 수정 권한 ---
        Dim isMine, isPwOk2
        isMine  = (Auth.MyId() > 0 And Auth.MyId() = artWriterId)
        isPwOk2 = (Session("PWOK_" & bdId & "_" & artId) & "" = "1")

        If Not (bdCanAdmin Or isMine) Then
            If artWriterId = 0 And artHasPw Then
                ' 비회원 글 → 비밀번호 확인 후 진입
                If Not isPwOk2 Then
                    Response.Redirect U("/board/PwCheck.asp") & "?bid=" & UE(bdId) & "&id=" & artId & "&act=modify"
                End If
            Else
                JsAlert "수정 권한이 없습니다."
            End If
        End If

        fTitle      = AV(artArr,  AVW_TITLE, 0)
        fContent    = AV(artArr,  AVW_CONTENT, 0)
        fCateId     = AVI(artArr, AVW_CID, 0)
        fTags       = AV(artArr,  AVW_TAGS, 0)
        fLinkUrl    = AV(artArr,  AVW_LINK, 0)
        fWriterName = AV(artArr,  AVW_WNAME, 0)
        fIsNotice   = AVB(artArr, AVW_NOTICE, 0)
        fIsSecret   = AVB(artArr, AVW_SECRET, 0)

    Else   ' reply
        fTitle   = "RE: " & AV(artArr, AVW_TITLE, 0)
        If Len(fTitle) > 200 Then fTitle = Left(fTitle, 200)
        fCateId  = AVI(artArr, AVW_CID, 0)
    End If
End If

'--- 분류 목록 ---
Dim cateArr, cateCnt
cateCnt = 0
If bdUseCategory Then
    cateArr = Board.Category_l(bdId)
    cateCnt = RowCnt(cateArr)
End If

'--- 기존 첨부 (수정 모드) ---
Dim fileArr, fileCnt
fileCnt = 0
If bdUseFile And wMode = "modify" Then
    fileArr = BoardFile.File_l(artId)
    fileCnt = RowCnt(fileArr)
End If

'--- 화면 텍스트 ---
Dim modeLabel
Select Case wMode
    Case "modify" : modeLabel = "수정"
    Case "reply"  : modeLabel = "답변 등록"
    Case Else     : modeLabel = "글쓰기"
End Select

Dim backUrl
backUrl = BoardBackUrl()

'--- 공지 지정은 게시판 관리자만 ---
Dim canSetNotice
canSetNotice = (bdUseNotice And bdCanAdmin)

g_pageTitle = bdName & " " & modeLabel
g_pageDesc  = bdName & " " & modeLabel
%>
