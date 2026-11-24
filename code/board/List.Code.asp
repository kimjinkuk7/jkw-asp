<!--#include file="Board.Code.asp"-->
<%
'==============================================================================
' code/board/List.Code.asp  -  게시판 목록 컨트롤러
'------------------------------------------------------------------------------
' 6개 게시판(공지·자유·자료실·갤러리·블로그·방명록)이 모두 이 파일을 쓴다.
' 화면 분기는 view/board/List.asp 에서 bdSkin 값으로만 이루어진다.
'==============================================================================

'--- 목록 열람 권한 ---
If Not bdCanList Then Auth.CheckLevel AVI(bdArr, BM_AUTH_LIST, 0)

'--- 파라미터 ---
Dim curPage, cateId, srchField, srchWord
curPage   = RequestUtil.ValueIntInit("page", 1)
cateId    = RequestUtil.ValueIntInit("cid", 0)
srchField = RequestUtil.ValueWhite("sf", "all,title,content,writer,tags", "all")
srchWord  = Left(RequestUtil.ValueInit("sw", ""), 50)

If curPage < 1 Then curPage = 1
If cateId > 0 Then
    If Not Board.IsValidCategory(bdId, cateId) Then cateId = 0
End If

'--- 총 건수 + 페이징 ---
Dim totalCount, pagingUrl
totalCount = Article.Article_c(bdId, cateId, srchField, srchWord)

pagingUrl = U("/board/List.asp") & "?bid=" & UE(bdId) & _
            "&cid=" & cateId & "&sf=" & UE(srchField) & "&sw=" & UE(srchWord) & "&page={P}"

With Paging
    .SetNowPage     = curPage
    .SetPageSize    = bdListCount
    .SetPageBlock   = bdPageBlock
    .SetResultCount = totalCount
    .SetPagingUrl   = pagingUrl
End With

'--- 목록 ---
Dim listArr, listCnt, listStartNo
listArr     = Article.Article_l(bdId, cateId, srchField, srchWord, Paging.Offset, bdListCount)
listCnt     = RowCnt(listArr)
listStartNo = Paging.StartNo

'--- 공지 (1페이지 & 검색 없을 때만) ---
Dim noticeArr, noticeCnt
noticeCnt = 0
If bdUseNotice And Paging.NowPage = 1 And StringUtil.IsBlank(srchWord) And cateId = 0 Then
    noticeArr = Article.Notice_l(bdId)
    noticeCnt = RowCnt(noticeArr)
End If

'--- 분류 탭 ---
Dim cateArr, cateCnt
cateCnt = 0
If bdUseCategory Then
    cateArr = Board.Category_l(bdId)
    cateCnt = RowCnt(cateArr)
End If

'--- 태그 클라우드 (블로그 스킨) ---
Dim tagArr, tagCnt
tagCnt = 0
If bdSkin = "blog" Then
    tagArr = Article.TagCloud_l(bdId, 20)
    tagCnt = RowCnt(tagArr)
End If

'--- 글쓰기 링크 ---
Dim writeUrl
writeUrl = U("/board/Write.asp") & "?bid=" & UE(bdId) & "&mode=write"

'--- 검색 중인지 ---
Dim isSearching
isSearching = (Not StringUtil.IsBlank(srchWord))

If isSearching Then g_pageTitle = bdName & " 검색 : " & srchWord
%>
