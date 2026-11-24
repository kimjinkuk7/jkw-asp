<%
'==============================================================================
' view/include/AdminSide.asp  -  관리자 사이드 메뉴
'   Code 파일에서 g_adminNav 에 활성 키를 지정한다.
'==============================================================================
%>
<aside class="admin-side">
    <h1>관리자</h1>
    <nav aria-label="관리 메뉴">
        <a href="<%= U("/admin/Default.asp")   %>" class="<%= IIfStr(g_adminNav = "dash",    "is-on", "") %>">대시보드</a>
        <a href="<%= U("/admin/Board.asp")     %>" class="<%= IIfStr(g_adminNav = "board",   "is-on", "") %>">게시판 관리</a>
        <a href="<%= U("/admin/Article.asp")   %>" class="<%= IIfStr(g_adminNav = "article", "is-on", "") %>">게시글 관리</a>
        <a href="<%= U("/admin/Member.asp")    %>" class="<%= IIfStr(g_adminNav = "member",  "is-on", "") %>">회원 관리</a>
        <a href="<%= U("/admin/Portfolio.asp") %>" class="<%= IIfStr(g_adminNav = "pf",      "is-on", "") %>">포트폴리오</a>
        <a href="<%= U("/admin/Career.asp")    %>" class="<%= IIfStr(g_adminNav = "career",  "is-on", "") %>">경력 · 학력</a>
        <a href="<%= U("/admin/Skill.asp")     %>" class="<%= IIfStr(g_adminNav = "skill",   "is-on", "") %>">기술 스택</a>
        <a href="<%= U("/admin/Contact.asp")   %>" class="<%= IIfStr(g_adminNav = "contact", "is-on", "") %>">
            문의 관리<% If admUnread > 0 Then %> <span class="badge badge-new"><%= admUnread %></span><% End If %>
        </a>
        <a href="<%= U("/admin/Stat.asp")      %>" class="<%= IIfStr(g_adminNav = "stat",    "is-on", "") %>">접속 통계</a>
        <a href="<%= U("/admin/Config.asp")    %>" class="<%= IIfStr(g_adminNav = "config",  "is-on", "") %>">사이트 설정</a>
        <a href="<%= U("/main/Default.asp")    %>" style="margin-top:12px;border-top:1px solid rgba(255,255,255,.12);padding-top:12px">← 사이트로</a>
    </nav>
</aside>
