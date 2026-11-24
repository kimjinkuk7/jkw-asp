<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Default.Code.asp  -  관리자 대시보드
'==============================================================================
g_adminNav  = "dash"
g_pageTitle = "관리자 대시보드"

'--- 요약 지표 (03_proc.sql 의 vw_dashboard_summary) ---
Dim dsArr, dsHelper
Set dsHelper = New CDBHelper
dsHelper.ConnectionString = ConnectionString(DBCONN_JKWCOM)
dsHelper.Sql = "SELECT member_total, member_today, article_total, article_today, " & _
               "       comment_total, contact_unread, visit_total, visit_today " & _
               "  FROM vw_dashboard_summary"
dsArr = dsHelper.ExecuteGetRows()
Set dsHelper = Nothing

Dim dsMemberT, dsMemberD, dsArticleT, dsArticleD, dsCommentT, dsUnread, dsVisitT, dsVisitD
If IsArray(dsArr) Then
    dsMemberT  = AVI(dsArr, 0, 0) : dsMemberD  = AVI(dsArr, 1, 0)
    dsArticleT = AVI(dsArr, 2, 0) : dsArticleD = AVI(dsArr, 3, 0)
    dsCommentT = AVI(dsArr, 4, 0) : dsUnread   = AVI(dsArr, 5, 0)
    dsVisitT   = AVI(dsArr, 6, 0) : dsVisitD   = AVI(dsArr, 7, 0)
Else
    dsMemberT = 0 : dsMemberD = 0 : dsArticleT = 0 : dsArticleD = 0
    dsCommentT = 0 : dsUnread = 0 : dsVisitT = 0 : dsVisitD = 0
End If

'--- 최근 게시글 / 댓글 / 문의 ---
Dim recArtArr, recArtCnt, recCmtArr, recCmtCnt, recCtArr, recCtCnt
recArtArr = Article.AdminArticle_l("", "", 0, 8) : recArtCnt = RowCnt(recArtArr)
recCmtArr = Comment.Recent_l(6)                  : recCmtCnt = RowCnt(recCmtArr)
recCtArr  = Contact.Contact_l("", False, 0, 6)   : recCtCnt  = RowCnt(recCtArr)

'--- 게시판별 글 수 ---
Dim bdListArr, bdListCnt
bdListArr = Board.Board_l(True)
bdListCnt = RowCnt(bdListArr)

'--- 최근 7일 방문 ---
Dim vsArr, vsCnt
vsArr = Visit.Stat_l(7)
vsCnt = RowCnt(vsArr)

'--- 그래프 최대값 ---
Dim vsMax, vi
vsMax = 1
For vi = 0 To vsCnt - 1
    If AVI(vsArr, 1, vi) > vsMax Then vsMax = AVI(vsArr, 1, vi)
Next
%>
