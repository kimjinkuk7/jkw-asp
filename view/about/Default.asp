<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/about/Default.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li>ABOUT</li>
            <li>소개</li>
        </ul>
        <h2>소개</h2>
        <p class="desc">어떤 개발자인지, 무엇을 해왔는지 정리했습니다.</p>
    </div>

    <!--#include file="../include/Lnb.asp"-->

    <section class="profile-head">
        <div class="photo" aria-hidden="true"></div>
        <div>
            <h2><%= HE(GLB_OWNER_NAME) %></h2>
            <p class="role"><%= HE(GLB_SITE_SUBTITLE) %></p>
            <div class="info-list">
                <div class="row"><span class="lbl">경력</span><span><%= HE(abYears) %>년</span></div>
                <div class="row"><span class="lbl">프로젝트</span><span><%= StringUtil.Comma(abProjectCnt) %>건</span></div>
<% If Len(GLB_OWNER_EMAIL) > 0 Then %>
                <div class="row"><span class="lbl">이메일</span><span><a href="mailto:<%= HE(GLB_OWNER_EMAIL) %>"><%= HE(GLB_OWNER_EMAIL) %></a></span></div>
<% End If %>
<% If Len(GLB_OWNER_GITHUB) > 0 Then %>
                <div class="row"><span class="lbl">GitHub</span><span><a href="<%= HE(GLB_OWNER_GITHUB) %>" target="_blank" rel="noopener noreferrer"><%= HE(GLB_OWNER_GITHUB) %></a></span></div>
<% End If %>
            </div>
        </div>
    </section>

    <section class="section">
        <div class="section-head"><h3>About Me</h3></div>
        <p style="white-space:pre-line;font-size:var(--fs-md);line-height:1.9"><%= HE(abIntro) %></p>
    </section>

<% If abWorkCnt > 0 Then %>
    <section class="section">
        <div class="section-head">
            <h3>경력</h3>
            <a class="more" href="<%= U("/about/Career.asp") %>">전체 보기 →</a>
        </div>
        <ul class="timeline">
<%      Dim wi
        For wi = 0 To abWorkCnt - 1
            If wi >= 3 Then Exit For %>
            <li>
                <span class="period"><%= HE(DateUtil.Period(AV(abWorkArr, CR_START, wi), AV(abWorkArr, CR_END, wi), AVB(abWorkArr, CR_CURRENT, wi))) %></span>
                <h4>
                    <%= AVH(abWorkArr, CR_TITLE, wi) %>
<%                  If AVB(abWorkArr, CR_CURRENT, wi) Then %><span class="badge badge-primary is-current">재직중</span><% End If %>
                </h4>
                <span class="position"><%= AVH(abWorkArr, CR_POSITION, wi) %></span>
            </li>
<%      Next %>
        </ul>
    </section>
<% End If %>

<% If abSkillCnt > 0 Then %>
    <section class="section">
        <div class="section-head">
            <h3>주요 기술</h3>
            <a class="more" href="<%= U("/about/Skill.asp") %>">전체 보기 →</a>
        </div>
        <div>
<%      Dim ski
        For ski = 0 To abSkillCnt - 1
            If ski >= 16 Then Exit For %>
            <span class="chip"><%= AVH(abSkillArr, SK_NAME, ski) %></span>
<%      Next %>
        </div>
    </section>
<% End If %>

    <div class="btn-row center mt-5">
        <a class="btn btn-primary btn-lg" href="<%= U("/about/Resume.asp") %>">이력서 전체 보기</a>
        <a class="btn btn-lg" href="<%= U("/contact/Default.asp") %>">문의하기</a>
    </div>

</div>

<!--#include file="../include/Footer.asp"-->
