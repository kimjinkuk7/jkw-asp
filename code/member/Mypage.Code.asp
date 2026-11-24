<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<%
'==============================================================================
' code/member/Mypage.Code.asp  -  마이페이지
'------------------------------------------------------------------------------
' tab : info(내 정보) | article(내 글) | comment(내 댓글) | password(비밀번호)
'==============================================================================
Auth.CheckLevel LV_MEMBER

g_pageTitle = "마이페이지"
g_pageDesc  = "내 정보와 활동 내역"
g_bodyClass = "page-member"

Dim myTab
myTab = LCase(RequestUtil.ValueWhite("tab", "info,article,comment,password,withdraw", "info"))

Dim meArr
meArr = Member.Member_v(GLB_MY_ID)
If Not IsArray(meArr) Then
    Auth.Logout
    JsAlertGo "회원 정보를 찾을 수 없습니다. 다시 로그인해 주세요.", U("/member/Login.asp")
End If

'--- 목록 (내 글 / 내 댓글) ---
Dim myPage, myPerPage, myTotal, myArr, myCnt
myPage    = RequestUtil.ValueIntInit("page", 1)
myPerPage = 10
myTotal   = 0
myCnt     = 0

If myTab = "article" Then
    myTotal = Member.MyArticle_c(GLB_MY_ID)
    With Paging
        .SetNowPage     = myPage
        .SetPageSize    = myPerPage
        .SetPageBlock   = DEF_PAGE_BLOCK
        .SetResultCount = myTotal
        .SetPagingUrl   = U("/member/Mypage.asp") & "?tab=article&page={P}"
    End With
    myArr = Member.MyArticle_l(GLB_MY_ID, Paging.Offset, myPerPage)
    myCnt = RowCnt(myArr)

ElseIf myTab = "comment" Then
    myArr = Member.MyComment_l(GLB_MY_ID, 0, 50)
    myCnt = RowCnt(myArr)
End If
%>
