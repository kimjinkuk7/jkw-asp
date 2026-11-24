<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/portfolio/Default.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li>PORTFOLIO</li>
        </ul>
        <h2>포트폴리오</h2>
        <p class="desc">진행한 프로젝트를 케이스 스터디 형태로 정리했습니다.</p>
    </div>

<% If pfCateCnt > 0 Then %>
    <nav class="cate-tabs" aria-label="카테고리">
        <a href="<%= U("/portfolio/Default.asp") %>" class="<%= IIfStr(pfCate = "", "is-on", "") %>">전체</a>
<%      Dim ci
        For ci = 0 To pfCateCnt - 1 %>
        <a href="<%= U("/portfolio/Default.asp") %>?cate=<%= UE(AV(pfCateArr, 0, ci)) %>"
           class="<%= IIfStr(pfCate = AV(pfCateArr, 0, ci), "is-on", "") %>"><%= AVH(pfCateArr, 0, ci) %> <%= AVI(pfCateArr, 1, ci) %></a>
<%      Next %>
    </nav>
<% End If %>

    <div class="board-head">
        <span class="count"><%= Paging.Summary() %></span>
    </div>

<% If pfCnt = 0 Then %>
    <div class="empty">등록된 프로젝트가 없습니다.</div>
<% Else %>
    <div class="pf-grid">
<%  Dim pi, pth
    For pi = 0 To pfCnt - 1
        pth = UpUrl(AV(pfArr, PF_THUMB, pi)) %>
        <a class="card pf-card" href="<%= U("/portfolio/View.asp") %>?id=<%= AVI(pfArr, PF_ID, pi) %>">
            <span class="thumb"<% If Len(pth) > 0 Then %> style="background-image:url('<%= HE(pth) %>')"<% End If %>>
<%          If Len(pth) = 0 Then %><span class="ph"><%= HE(Left(AV(pfArr, PF_TITLE, pi), 2)) %></span><% End If %>
            </span>
            <span class="body">
                <span class="cate"><%= AVH(pfArr, PF_CATEGORY, pi) %></span>
                <h3><%= AVH(pfArr, PF_TITLE, pi) %></h3>
                <p class="summary"><%= AVH(pfArr, PF_SUMMARY, pi) %></p>
                <span class="period">
                    <%= HE(DateUtil.Period(AV(pfArr, PF_START, pi), AV(pfArr, PF_END, pi), AVB(pfArr, PF_ONGOING, pi))) %>
<%                  If Len(AV(pfArr, PF_CLIENT, pi)) > 0 Then %> · <%= AVH(pfArr, PF_CLIENT, pi) %><% End If %>
                </span>
            </span>
        </a>
<%  Next %>
    </div>
<% End If %>

    <%= Paging.Render() %>

</div>

<!--#include file="../include/Footer.asp"-->
