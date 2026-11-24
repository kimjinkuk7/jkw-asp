<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/board/Board.Class.asp"-->
<!--#include file="../../model/board/Article.Class.asp"-->
<!--#include file="../../model/board/Comment.Class.asp"-->
<!--#include file="../../model/board/BoardFile.Class.asp"-->
<%
'==============================================================================
' code/board/Board.Code.asp  -  게시판 모듈 공통
'------------------------------------------------------------------------------
' 모든 board Code 파일이 이 파일을 첫 줄에 include 한다.
' 여기서 bid 를 검증하고 게시판 설정을 전역 변수(bd*)로 펼쳐놓는다.
'==============================================================================

'--- board_id ---------------------------------------------------------------
' ★ 반드시 QueryString 을 먼저 읽는다.
'   Classic ASP 는 Request.Form 을 한 번이라도 읽으면 Request.BinaryRead 가
'   실패하므로(파일 업로드 불가), multipart 요청에서는 Form 을 건드리면 안 된다.
'   그래서 파일 업로드 폼의 action 에는 항상 ?bid=... 를 붙여둔다.
Dim bdId, bdArr
bdId = LCase(Trim(Request.QueryString("bid") & ""))
If Len(bdId) = 0 Then
    If Not RequestUtil.IsMultipart() Then bdId = LCase(RequestUtil.ValueInit("bid", ""))
End If

' 형식 검증 : 영소문자로 시작하는 영숫자/_ 2~20자
If Not StringUtil.RegTest("^[a-z][a-z0-9_]{1,19}$", bdId) Then
    Go404 "게시판을 찾을 수 없습니다."
End If

bdArr = Board.Board_v(bdId)
If Not IsArray(bdArr) Then
    Go404 "존재하지 않거나 사용 중지된 게시판입니다."
End If

'--- 게시판 설정 펼치기 -------------------------------------------------------
Dim bdName, bdDesc, bdSkin, bdListCount, bdPageBlock, bdGalCol
Dim bdUseComment, bdUseFile, bdUseCategory, bdUseSecret, bdUseReply, bdUseEditor, bdUseNotice
Dim bdFileCnt, bdFileSize, bdFileExt

bdName        = AV(bdArr,  BM_NAME, 0)
bdDesc        = AV(bdArr,  BM_DESC, 0)
bdSkin        = AV(bdArr,  BM_SKIN, 0)
bdListCount   = AVI(bdArr, BM_LIST_COUNT, 0)
bdPageBlock   = AVI(bdArr, BM_PAGE_BLOCK, 0)
bdGalCol      = AVI(bdArr, BM_GAL_COL, 0)

bdUseComment  = AVB(bdArr, BM_USE_COMMENT, 0)
bdUseFile     = AVB(bdArr, BM_USE_FILE, 0)
bdUseCategory = AVB(bdArr, BM_USE_CATEGORY, 0)
bdUseSecret   = AVB(bdArr, BM_USE_SECRET, 0)
bdUseReply    = AVB(bdArr, BM_USE_REPLY, 0)
bdUseEditor   = AVB(bdArr, BM_USE_EDITOR, 0)
bdUseNotice   = AVB(bdArr, BM_USE_NOTICE, 0)

bdFileCnt     = AVI(bdArr, BM_FILE_CNT, 0)
bdFileSize    = AVI(bdArr, BM_FILE_SIZE, 0)
bdFileExt     = AV(bdArr,  BM_FILE_EXT, 0)

'--- 권한 플래그 -------------------------------------------------------------
Dim bdCanList, bdCanRead, bdCanWrite, bdCanComment, bdCanAdmin
bdCanList    = Board.CanList(bdArr)
bdCanRead    = Board.CanRead(bdArr)
bdCanWrite   = Board.CanWrite(bdArr)
bdCanComment = Board.CanComment(bdArr)
bdCanAdmin   = Board.CanAdmin(bdArr)

'--- 현재 메뉴 하이라이트 -----------------------------------------------------
g_menuCode  = MenuCodeByBid(bdId)
g_pageTitle = bdName
g_pageDesc  = bdDesc
g_bodyClass = "page-board board-" & bdSkin


'##############################################################################
' 게시판 모듈 공통 함수
'##############################################################################

'<Desc>board_id → tb_menu.menu_code 매핑 (GNB 하이라이트용)</Desc>
Function MenuCodeByBid(ByVal pBid)
    Select Case pBid
        Case "notice"  : MenuCodeByBid = "comm_notice"
        Case "free"    : MenuCodeByBid = "comm_free"
        Case "data"    : MenuCodeByBid = "comm_data"
        Case "gallery" : MenuCodeByBid = "comm_gallery"
        Case "blog"    : MenuCodeByBid = "blog"
        Case "guest"   : MenuCodeByBid = "guestbook"
        Case Else      : MenuCodeByBid = "comm_" & pBid
    End Select
End Function

'<Desc>목록으로 돌아가는 URL (페이지/검색 상태 유지)</Desc>
Function BoardBackUrl()
    BoardBackUrl = U("/board/List.asp") & "?bid=" & UE(bdId) & ListStateQs()
End Function

'<Desc>비회원 작성자명 / 비밀번호 처리용 : 로그인 상태면 세션값 사용</Desc>
Function WriterName(ByVal pInputName)
    If Auth.IsLogin() Then
        WriterName = Auth.DisplayName()
    Else
        WriterName = Left(Trim(pInputName), 20)
    End If
End Function

'<Desc>글/댓글 수정·삭제 가능 여부 (관리자 / 본인 / 비회원 비밀번호 보유)</Desc>
Function CanEditArticle(ByVal pWriterId, ByVal pHasPw)
    CanEditArticle = False
    If bdCanAdmin Then CanEditArticle = True : Exit Function
    If Auth.MyId() > 0 And StringUtil.ToInt(pWriterId, 0) = Auth.MyId() Then CanEditArticle = True : Exit Function
    If StringUtil.ToInt(pWriterId, 0) = 0 And StringUtil.ToBool(pHasPw) Then CanEditArticle = True
End Function
%>
