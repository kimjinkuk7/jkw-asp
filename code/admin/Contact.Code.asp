<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Contact.Code.asp  -  문의 관리 (목록 + 상세/답변)
'==============================================================================
g_adminNav  = "contact"
g_pageTitle = "문의 관리"

Dim ctId, ctArr, ctIsView
ctId     = RequestUtil.ValueIntInit("id", 0)
ctIsView = False

If ctId > 0 Then
    ctArr = Contact.Contact_v(ctId)
    If IsArray(ctArr) Then
        ctIsView = True
        '--- 열람 시 자동으로 확인 처리 ---
        If Not AVB(ctArr, CT_ISREAD, 0) Then Contact.Read_u ctId
    Else
        JsAlertGo "존재하지 않는 문의입니다.", U("/admin/Contact.asp")
    End If
End If

'--- 목록 ---
Dim ctWord, ctUnread, ctPage, ctPerPage, ctTotal, ctListArr, ctListCnt
ctWord    = Left(RequestUtil.ValueInit("sw", ""), 50)
ctUnread  = (RequestUtil.ValueInit("unread", "") = "1")
ctPage    = RequestUtil.ValueIntInit("page", 1)
ctPerPage = 20

ctTotal = Contact.Contact_c(ctWord, ctUnread)

With Paging
    .SetNowPage     = ctPage
    .SetPageSize    = ctPerPage
    .SetPageBlock   = DEF_PAGE_BLOCK
    .SetResultCount = ctTotal
    .SetPagingUrl   = U("/admin/Contact.asp") & "?sw=" & UE(ctWord) & "&unread=" & IIfStr(ctUnread, "1", "") & "&page={P}"
End With

ctListArr = Contact.Contact_l(ctWord, ctUnread, Paging.Offset, ctPerPage)
ctListCnt = RowCnt(ctListArr)
%>
