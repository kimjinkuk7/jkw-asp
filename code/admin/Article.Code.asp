<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Article.Code.asp  -  게시글 통합 관리
'==============================================================================
g_adminNav  = "article"
g_pageTitle = "게시글 관리"

Dim aaBid, aaWord, aaPage, aaTotal, aaPerPage
aaBid     = LCase(Trim(RequestUtil.ValueInit("bid", "")))
aaWord    = Left(RequestUtil.ValueInit("sw", ""), 50)
aaPage    = RequestUtil.ValueIntInit("page", 1)
aaPerPage = 20

If Len(aaBid) > 0 Then
    If Not StringUtil.RegTest("^[a-z][a-z0-9_]{1,19}$", aaBid) Then aaBid = ""
End If

aaTotal = Article.AdminArticle_c(aaBid, aaWord)

With Paging
    .SetNowPage     = aaPage
    .SetPageSize    = aaPerPage
    .SetPageBlock   = DEF_PAGE_BLOCK
    .SetResultCount = aaTotal
    .SetPagingUrl   = U("/admin/Article.asp") & "?bid=" & UE(aaBid) & "&sw=" & UE(aaWord) & "&page={P}"
End With

Dim aaArr, aaCnt
aaArr = Article.AdminArticle_l(aaBid, aaWord, Paging.Offset, aaPerPage)
aaCnt = RowCnt(aaArr)

Dim aaBoards, aaBoardCnt
aaBoards   = Board.Board_l(True)
aaBoardCnt = RowCnt(aaBoards)
%>
