<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/portfolio/Portfolio.Class.asp"-->
<%
'==============================================================================
' code/portfolio/Default.Code.asp  -  포트폴리오 목록
'==============================================================================
g_menuCode  = "portfolio"
g_pageTitle = "포트폴리오"
g_pageDesc  = GLB_OWNER_NAME & "가 수행한 프로젝트 목록"
g_bodyClass = "page-portfolio"

'--- 파라미터 ---
Dim pfCate, pfPage
pfCate = Left(RequestUtil.ValueInit("cate", ""), 50)
pfPage = RequestUtil.ValueIntInit("page", 1)

'--- 카테고리 탭 ---
Dim pfCateArr, pfCateCnt
pfCateArr = Portfolio.Category_l()
pfCateCnt = RowCnt(pfCateArr)

'--- 유효하지 않은 카테고리는 전체로 ---
If Len(pfCate) > 0 Then
    Dim ok, k
    ok = False
    For k = 0 To pfCateCnt - 1
        If AV(pfCateArr, 0, k) = pfCate Then ok = True
    Next
    If Not ok Then pfCate = ""
End If

'--- 목록 + 페이징 ---
Dim pfTotal, pfPerPage
pfPerPage = 9
pfTotal   = Portfolio.Portfolio_c(pfCate)

With Paging
    .SetNowPage     = pfPage
    .SetPageSize    = pfPerPage
    .SetPageBlock   = DEF_PAGE_BLOCK
    .SetResultCount = pfTotal
    .SetPagingUrl   = U("/portfolio/Default.asp") & "?cate=" & UE(pfCate) & "&page={P}"
End With

Dim pfArr, pfCnt
pfArr = Portfolio.Portfolio_l(pfCate, Paging.Offset, pfPerPage)
pfCnt = RowCnt(pfArr)
%>
