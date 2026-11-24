<%
'==============================================================================
' view/include/Footer.asp  -  공통 하단 (</main> ~ </html>)
'==============================================================================
Dim ftCnt, ftToday, ftTotal
ftCnt = Visit.Counter_v()
If IsArray(ftCnt) Then
    ftToday = StringUtil.Comma(AVI(ftCnt, 0, 0))
    ftTotal = StringUtil.Comma(AVI(ftCnt, 1, 0))
Else
    ftToday = "0" : ftTotal = "0"
End If
%>
</main>

<footer id="footer">
    <div class="wrap">

        <div class="footer-top">
            <div class="footer-brand">
                <span class="logo"><%= HE(GLB_SITE_TITLE) %><b class="dot">.</b></span>
                <p><%= HE(GLB_SITE_DESC) %></p>
<% If Len(GLB_OWNER_EMAIL) > 0 Then %>
                <p><a href="mailto:<%= HE(GLB_OWNER_EMAIL) %>"><%= HE(GLB_OWNER_EMAIL) %></a></p>
<% End If %>
            </div>

            <nav class="footer-links" aria-label="하단 메뉴">
                <a href="<%= U("/about/Default.asp") %>">소개</a>
                <a href="<%= U("/portfolio/Default.asp") %>">포트폴리오</a>
                <a href="<%= U("/board/List.asp") %>?bid=blog">블로그</a>
                <a href="<%= U("/board/List.asp") %>?bid=notice">공지사항</a>
                <a href="<%= U("/board/List.asp") %>?bid=guest">방명록</a>
                <a href="<%= U("/contact/Default.asp") %>">문의하기</a>
<% If Len(GLB_OWNER_GITHUB) > 0 Then %>
                <a href="<%= HE(GLB_OWNER_GITHUB) %>" target="_blank" rel="noopener noreferrer">GitHub</a>
<% End If %>
            </nav>
        </div>

        <div class="footer-bottom">
            <span><%= GLB_COPYRIGHT %></span>
            <span class="footer-stat">
                <span>Today <strong><%= ftToday %></strong></span>
                <span>Total <strong><%= ftTotal %></strong></span>
            </span>
        </div>

    </div>
</footer>

<button type="button" id="btn-top" class="btn btn-ghost"
        style="display:none;position:fixed;right:16px;bottom:16px;z-index:90;width:44px;height:44px;border-radius:50%;background:var(--c-text);color:#fff;box-shadow:var(--shadow-lg);"
        aria-label="맨 위로">↑</button>

<script src="<%= APP_ROOT %>/assets/js/common.js?v=<%= HE(GLB_ASSET_VER) %>"></script>
<%= g_extraScript %>
</body>
</html>
<%
'--- 요청 종료 : 공유 커넥션 정리 ---
DBDisconnect
%>
