<%
'==============================================================================
' view/include/Lnb.asp  -  하위 메뉴 탭
'------------------------------------------------------------------------------
' Code 파일에서 g_lnbParent 에 1depth menu_code 를 지정하면 자동으로 그려진다.
'   g_lnbParent = "about"      → 소개 / 경력·학력 / 기술스택 / 이력서
'   g_lnbParent = "community"  → 공지사항 / 자유게시판 / 자료실 / 갤러리
'==============================================================================
If Len(g_lnbParent) > 0 Then
    Dim lnbArr, lnbCnt, lnbI
    lnbArr = Menu.SubMenu_l(g_lnbParent, GLB_MY_LEVEL)
    lnbCnt = RowCnt(lnbArr)

    If lnbCnt > 0 Then
%>
<nav class="cate-tabs" aria-label="하위 메뉴">
<%      For lnbI = 0 To lnbCnt - 1 %>
    <a href="<%= U(AV(lnbArr, MNU_URL, lnbI)) %>"
       class="<%= IIfStr(g_menuCode = AV(lnbArr, MNU_CODE, lnbI), "is-on", "") %>"><%= AVH(lnbArr, MNU_NAME, lnbI) %></a>
<%      Next %>
</nav>
<%
    End If
End If
%>
