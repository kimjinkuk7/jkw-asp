<%
'==============================================================================
' view/include/Gnb.asp  -  전역 네비게이션
'------------------------------------------------------------------------------
' tb_menu 를 그대로 렌더링한다. 메뉴 추가 = DB INSERT 1건, 코드 수정 0줄.
' 모바일에서는 슬라이드 패널, 1024px 이상에서는 가로 GNB + 드롭다운. (CSS 로 전환)
'------------------------------------------------------------------------------
' ★ 이 파일은 Header.asp 를 통해 "모든 페이지"에 include 된다.
'   따라서 여기서 선언하는 변수는 전부 gnb 접두어를 붙인다.
'   짧은 이름(gi, gj 등)을 쓰면 각 화면의 루프 변수와 충돌해
'   VBScript 오류 800a0411(다시 정의된 이름)이 발생한다.
'==============================================================================
Dim gnbArr, gnbCnt, gnbI, gnbJ
Dim gnbChildStart, gnbChildEnd, gnbCurId, gnbIsOn, gnbHasChild

gnbArr = Menu.Menu_l(GLB_MY_LEVEL)
gnbCnt = RowCnt(gnbArr)
%>
<nav id="gnb" aria-label="주 메뉴">
    <ul>
<%
gnbI = 0
Do While gnbI < gnbCnt

    If AVI(gnbArr, MNU_DEPTH, gnbI) = 1 Then

        gnbCurId = AVI(gnbArr, MNU_ID, gnbI)

        '--- 하위 메뉴 구간 찾기 ---
        gnbChildStart = gnbI + 1
        gnbChildEnd   = gnbI
        gnbJ = gnbI + 1
        Do While gnbJ < gnbCnt
            If AVI(gnbArr, MNU_DEPTH, gnbJ) = 2 And AVI(gnbArr, MNU_PARENT, gnbJ) = gnbCurId Then
                gnbChildEnd = gnbJ
                gnbJ = gnbJ + 1
            Else
                Exit Do
            End If
        Loop
        gnbHasChild = (gnbChildEnd >= gnbChildStart)

        '--- 현재 메뉴 여부 ---
        gnbIsOn = (g_menuCode = AV(gnbArr, MNU_CODE, gnbI))
        If Not gnbIsOn And gnbHasChild Then
            For gnbJ = gnbChildStart To gnbChildEnd
                If g_menuCode = AV(gnbArr, MNU_CODE, gnbJ) Then gnbIsOn = True
            Next
        End If
%>
        <li class="<%= IIfStr(gnbIsOn, "is-on", "") %>">
<%      If gnbHasChild Then %>
            <span class="gnb-label" role="button" tabindex="0"><%= AVH(gnbArr, MNU_NAME, gnbI) %></span>
            <div class="sub-menu">
<%          For gnbJ = gnbChildStart To gnbChildEnd %>
                <a href="<%= U(AV(gnbArr, MNU_URL, gnbJ)) %>"
                   class="<%= IIfStr(g_menuCode = AV(gnbArr, MNU_CODE, gnbJ), "is-on", "") %>"><%= AVH(gnbArr, MNU_NAME, gnbJ) %></a>
<%          Next %>
            </div>
<%      Else %>
            <a href="<%= U(AV(gnbArr, MNU_URL, gnbI)) %>"><%= AVH(gnbArr, MNU_NAME, gnbI) %></a>
<%      End If %>
        </li>
<%
        gnbI = gnbChildEnd + 1
    Else
        gnbI = gnbI + 1
    End If
Loop
%>
    </ul>

    <!-- 모바일 전용 계정 영역 -->
    <div class="gnb-account">
<% If GLB_IS_LOGIN Then %>
        <a class="btn btn-sm" href="<%= U("/member/Mypage.asp") %>"><%= HE(GLB_MY_NAME) %> 님</a>
    <% If GLB_MY_LEVEL >= LV_ADMIN Then %>
        <a class="btn btn-sm" href="<%= U("/admin/Default.asp") %>">관리자</a>
    <% End If %>
        <a class="btn btn-sm" href="<%= U("/member/Logout.asp") %>">로그아웃</a>
<% Else %>
        <a class="btn btn-sm btn-primary" href="<%= U("/member/Login.asp") %>">로그인</a>
        <a class="btn btn-sm" href="<%= U("/member/Join.asp") %>">회원가입</a>
<% End If %>
    </div>
</nav>
