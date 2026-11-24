<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/main/Default.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<!-- ================= HERO ================= -->
<section class="hero">
    <div class="wrap">
        <span class="eyebrow">Backend Developer</span>
        <h1><%= HE(heroHeadline) %></h1>
        <p class="lead"><%= HE(heroSub) %></p>

        <div class="btn-row">
            <a class="btn btn-primary btn-lg" href="<%= U("/portfolio/Default.asp") %>">포트폴리오 보기</a>
            <a class="btn btn-lg" href="<%= U("/about/Resume.asp") %>">이력서</a>
            <a class="btn btn-lg btn-ghost" href="<%= U("/contact/Default.asp") %>">문의하기</a>
        </div>

<% If skillCnt > 0 Then %>
        <div class="stack">
<%      Dim si
        For si = 0 To skillCnt - 1
            If si >= mainSkillMax Then Exit For %>
            <span class="chip"><%= AVH(skillArr, SK_NAME, si) %></span>
<%      Next %>
        </div>
<% End If %>

        <div class="stat-row">
            <div class="stat"><span class="num"><%= HE(statYears) %></span><span class="lbl">경력 (년)</span></div>
            <div class="stat"><span class="num"><%= StringUtil.Comma(statProject) %></span><span class="lbl">프로젝트</span></div>
            <div class="stat"><span class="num"><%= StringUtil.Comma(statPost) %></span><span class="lbl">블로그 글</span></div>
            <div class="stat"><span class="num"><%= StringUtil.Comma(statVisit) %></span><span class="lbl">누적 방문</span></div>
        </div>
    </div>
</section>

<!-- ================= 대표 프로젝트 ================= -->
<section class="section">
    <div class="wrap">
        <div class="section-head">
            <h3>대표 프로젝트</h3>
            <a class="more" href="<%= U("/portfolio/Default.asp") %>">전체 보기 →</a>
        </div>

<% If mainPfCnt = 0 Then %>
        <div class="empty">등록된 프로젝트가 없습니다.</div>
<% Else %>
        <div class="pf-grid main-pf-list">
<%      Dim pi
        For pi = 0 To mainPfCnt - 1 %>
            <a class="card pf-card" href="<%= U("/portfolio/View.asp") %>?id=<%= AVI(mainPfArr, PF_ID, pi) %>">
                <span class="thumb"<% If Len(AV(mainPfArr, PF_THUMB, pi)) > 0 Then %> style="background-image:url('<%= HE(UpUrl(AV(mainPfArr, PF_THUMB, pi))) %>')"<% End If %>>
<%              If Len(AV(mainPfArr, PF_THUMB, pi)) = 0 Then %><span class="ph"><%= HE(Left(AV(mainPfArr, PF_TITLE, pi), 2)) %></span><% End If %>
                </span>
                <span class="body">
                    <span class="cate"><%= AVH(mainPfArr, PF_CATEGORY, pi) %></span>
                    <h3><%= AVH(mainPfArr, PF_TITLE, pi) %></h3>
                    <p class="summary"><%= AVH(mainPfArr, PF_SUMMARY, pi) %></p>
                    <span class="period"><%= HE(DateUtil.Period(AV(mainPfArr, PF_START, pi), AV(mainPfArr, PF_END, pi), AVB(mainPfArr, PF_ONGOING, pi))) %></span>
                </span>
            </a>
<%      Next %>
        </div>
<% End If %>
    </div>
</section>

<!-- ================= 최근 글 ================= -->
<section class="section">
    <div class="wrap">
        <div class="grid" style="grid-template-columns:1fr;gap:var(--sp-6)">

            <div>
                <div class="section-head">
                    <h3>최근 블로그</h3>
                    <a class="more" href="<%= U("/board/List.asp") %>?bid=blog">더 보기 →</a>
                </div>
<% If blogCnt = 0 Then %>
                <div class="empty">등록된 글이 없습니다.</div>
<% Else %>
                <ul class="main-post-list">
<%      Dim bi
        For bi = 0 To blogCnt - 1 %>
                    <li>
                        <a href="<%= BoardViewUrl("blog", AVI(blogArr, ART_ID, bi)) %>">
                            <span class="tit"><%= AVH(blogArr, ART_TITLE, bi) %></span>
                            <span class="date"><%= HE(DateUtil.Fmt(AV(blogArr, ART_REGDT, bi), "ymd")) %></span>
                        </a>
                    </li>
<%      Next %>
                </ul>
<% End If %>
            </div>

            <div>
                <div class="section-head">
                    <h3>공지사항</h3>
                    <a class="more" href="<%= U("/board/List.asp") %>?bid=notice">더 보기 →</a>
                </div>
<% If noticeCnt = 0 Then %>
                <div class="empty">등록된 공지가 없습니다.</div>
<% Else %>
                <ul class="main-post-list">
<%      Dim ni
        For ni = 0 To noticeCnt - 1 %>
                    <li>
                        <a href="<%= BoardViewUrl("notice", AVI(noticeArr, ART_ID, ni)) %>">
                            <span class="tit">
                                <%= AVH(noticeArr, ART_TITLE, ni) %>
<%                              If DateUtil.IsNew(AV(noticeArr, ART_REGDT, ni), 48) Then %><span class="badge badge-new">N</span><% End If %>
                            </span>
                            <span class="date"><%= HE(DateUtil.Fmt(AV(noticeArr, ART_REGDT, ni), "ymd")) %></span>
                        </a>
                    </li>
<%      Next %>
                </ul>
<% End If %>
            </div>

        </div>
    </div>
</section>

<!--#include file="../include/Footer.asp"-->
