<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Member.Code.asp  -  회원 관리
'==============================================================================
g_adminNav  = "member"
g_pageTitle = "회원 관리"

Dim amWord, amLevel, amPage, amPerPage, amTotal
amWord    = Left(RequestUtil.ValueInit("sw", ""), 50)
amLevel   = RequestUtil.ValueIntInit("lv", -1)
amPage    = RequestUtil.ValueIntInit("page", 1)
amPerPage = 20

If InStr(",0,1,3,5,9,", "," & amLevel & ",") = 0 Then amLevel = -1

amTotal = Member.Member_c(amWord, amLevel)

With Paging
    .SetNowPage     = amPage
    .SetPageSize    = amPerPage
    .SetPageBlock   = DEF_PAGE_BLOCK
    .SetResultCount = amTotal
    .SetPagingUrl   = U("/admin/Member.asp") & "?sw=" & UE(amWord) & "&lv=" & amLevel & "&page={P}"
End With

Dim amArr, amCnt
amArr = Member.Member_l(amWord, amLevel, Paging.Offset, amPerPage)
amCnt = RowCnt(amArr)

Function LvName(ByVal lv)
    Select Case CLng(lv)
        Case 9 : LvName = "최고관리자"
        Case 5 : LvName = "운영자"
        Case 3 : LvName = "일반회원"
        Case 0 : LvName = "정지"
        Case Else : LvName = "Lv." & lv
    End Select
End Function
%>
