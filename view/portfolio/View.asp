<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/portfolio/View.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap wrap-narrow">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li><a href="<%= U("/portfolio/Default.asp") %>">PORTFOLIO</a></li>
            <li><%= HE(StringUtil.Cut(pvTitle, 20)) %></li>
        </ul>
    </div>

    <article>
        <header class="pf-view-head">
<% If Len(pvCate) > 0 Then %>
            <span class="badge badge-primary"><%= HE(pvCate) %></span>
<% End If %>
            <h2><%= HE(pvTitle) %></h2>
<% If Len(pvSubtitle) > 0 Then %>
            <p class="subtitle"><%= HE(pvSubtitle) %></p>
<% End If %>
        </header>

<% If Len(pvThumb) > 0 Then %>
        <img src="<%= HE(pvThumb) %>" alt="<%= HE(pvTitle) %> 대표 이미지" style="width:100%;border-radius:var(--radius);border:1px solid var(--c-border)">
<% End If %>

<%'--- 프로젝트 개요 ---%>
        <div class="pf-info">
            <div class="row"><span class="lbl">기간</span><span><%= HE(pvPeriod) %></span></div>
<% If Len(pvClient) > 0 Then %>
            <div class="row"><span class="lbl">발주/소속</span><span><%= HE(pvClient) %></span></div>
<% End If %>
<% If Len(pvRole) > 0 Then %>
            <div class="row"><span class="lbl">담당 역할</span><span><%= HE(pvRole) %></span></div>
<% End If %>
<% If Len(pvTeam) > 0 Then %>
            <div class="row"><span class="lbl">규모</span><span><%= HE(pvTeam) %></span></div>
<% End If %>
<% If pvSkillCnt > 0 Then %>
            <div class="row">
                <span class="lbl">기술</span>
                <span>
<%              Dim si
                For si = 0 To pvSkillCnt - 1 %>
                    <span class="chip"<% If Len(AV(pvSkillArr, 1, si)) > 0 Then %> style="border-left:3px solid <%= HE(AV(pvSkillArr, 1, si)) %>"<% End If %>><%= AVH(pvSkillArr, 0, si) %></span>
<%              Next %>
                </span>
            </div>
<% End If %>
            <div class="row"><span class="lbl">조회</span><span><%= StringUtil.Comma(pvHit) %></span></div>
        </div>

<% If Len(pvSummary) > 0 Then %>
        <p style="font-size:var(--fs-md);color:var(--c-text-sub);line-height:1.9"><%= HE(pvSummary) %></p>
<% End If %>

<% If Len(pvSiteUrl) > 0 Or Len(pvRepoUrl) > 0 Then %>
        <div class="btn-row mt-4">
<%      If Len(pvSiteUrl) > 0 Then %>
            <a class="btn btn-primary" href="<%= HE(pvSiteUrl) %>" target="_blank" rel="noopener noreferrer">사이트 보기 ↗</a>
<%      End If %>
<%      If Len(pvRepoUrl) > 0 Then %>
            <a class="btn" href="<%= HE(pvRepoUrl) %>" target="_blank" rel="noopener noreferrer">저장소 보기 ↗</a>
<%      End If %>
        </div>
<% End If %>

<% If Len(pvContent) > 0 Then %>
        <div class="view-body mt-5" style="border-top:1px solid var(--c-border)">
            <%= StringUtil.SafeHtml(pvContent) %>
        </div>
<% End If %>
    </article>

    <div class="btn-row between mt-5">
        <a class="btn" href="<%= U("/portfolio/Default.asp") %>">목록</a>
        <a class="btn btn-primary" href="<%= U("/contact/Default.asp") %>">이런 작업 문의하기</a>
    </div>

    <ul class="view-nav">
<% If IsArray(pvPrev) Then %>
        <li>
            <span class="dir">이전</span>
            <a class="tit" href="<%= U("/portfolio/View.asp") %>?id=<%= AVI(pvPrev, 0, 0) %>"><%= AVH(pvPrev, 1, 0) %></a>
        </li>
<% End If %>
<% If IsArray(pvNext) Then %>
        <li>
            <span class="dir">다음</span>
            <a class="tit" href="<%= U("/portfolio/View.asp") %>?id=<%= AVI(pvNext, 0, 0) %>"><%= AVH(pvNext, 1, 0) %></a>
        </li>
<% End If %>
    </ul>

</div>

<!--#include file="../include/Footer.asp"-->
