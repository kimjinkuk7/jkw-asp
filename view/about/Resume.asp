<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/about/Resume.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap wrap-narrow">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li>ABOUT</li>
            <li>이력서</li>
        </ul>
        <h2>이력서</h2>
        <p class="desc">브라우저 인쇄(Ctrl+P) 시 이력서만 깔끔하게 출력됩니다.</p>
    </div>

    <!--#include file="../include/Lnb.asp"-->

    <div class="btn-row right mb-4">
        <button type="button" class="btn btn-sm" onclick="window.print()">인쇄 / PDF 저장</button>
    </div>

    <article class="card" style="padding:var(--sp-6)">

<%'--- 헤더 ---%>
        <header style="border-bottom:2px solid var(--c-text);padding-bottom:var(--sp-4);margin-bottom:var(--sp-5)">
            <h2 style="font-size:var(--fs-2xl)"><%= HE(GLB_OWNER_NAME) %></h2>
            <p class="role" style="color:var(--c-primary);font-weight:700"><%= HE(GLB_SITE_SUBTITLE) %></p>
            <div class="info-list mt-3">
                <div class="row"><span class="lbl">총 경력</span><span><%= HE(rsYears) %>년</span></div>
<% If Len(GLB_OWNER_EMAIL) > 0 Then %>
                <div class="row"><span class="lbl">이메일</span><span><%= HE(GLB_OWNER_EMAIL) %></span></div>
<% End If %>
<% If Len(GLB_OWNER_GITHUB) > 0 Then %>
                <div class="row"><span class="lbl">GitHub</span><span><%= HE(GLB_OWNER_GITHUB) %></span></div>
<% End If %>
                <div class="row"><span class="lbl">홈페이지</span><span><%= HE(GLB_SITE_URL) %></span></div>
            </div>
        </header>

<%'--- 요약 ---%>
        <section class="mb-5">
            <h3 style="font-size:var(--fs-md);margin-bottom:var(--sp-2)">SUMMARY</h3>
            <p style="white-space:pre-line;color:var(--c-text-sub)"><%= HE(rsIntro) %></p>
        </section>

<%'--- 경력 ---%>
<% If rsWorkCnt > 0 Then %>
        <section class="mb-5">
            <h3 style="font-size:var(--fs-md);margin-bottom:var(--sp-3)">EXPERIENCE</h3>
            <ul class="timeline">
<%          Dim wi
            For wi = 0 To rsWorkCnt - 1 %>
                <li>
                    <span class="period"><%= HE(DateUtil.Period(AV(rsWork, CR_START, wi), AV(rsWork, CR_END, wi), AVB(rsWork, CR_CURRENT, wi))) %></span>
                    <h4><%= AVH(rsWork, CR_TITLE, wi) %></h4>
                    <span class="position"><%= AVH(rsWork, CR_POSITION, wi) %></span>
<%                  If Len(AV(rsWork, CR_DESC, wi)) > 0 Then %>
                    <div class="desc"><%= AVH(rsWork, CR_DESC, wi) %></div>
<%                  End If %>
                </li>
<%          Next %>
            </ul>
        </section>
<% End If %>

<%'--- 주요 프로젝트 ---%>
<% If rsPfCnt > 0 Then %>
        <section class="mb-5">
            <h3 style="font-size:var(--fs-md);margin-bottom:var(--sp-3)">PROJECTS</h3>
            <ul class="timeline">
<%          Dim pi
            For pi = 0 To rsPfCnt - 1 %>
                <li>
                    <span class="period"><%= HE(DateUtil.Period(AV(rsPf, PF_START, pi), AV(rsPf, PF_END, pi), AVB(rsPf, PF_ONGOING, pi))) %></span>
                    <h4><%= AVH(rsPf, PF_TITLE, pi) %></h4>
                    <span class="position"><%= AVH(rsPf, PF_ROLE, pi) %><% If Len(AV(rsPf, PF_CLIENT, pi)) > 0 Then %> · <%= AVH(rsPf, PF_CLIENT, pi) %><% End If %></span>
                    <div class="desc"><%= AVH(rsPf, PF_SUMMARY, pi) %></div>
                </li>
<%          Next %>
            </ul>
        </section>
<% End If %>

<%'--- 기술 ---%>
<% If rsSkillGrpCnt > 0 Then %>
        <section class="mb-5">
            <h3 style="font-size:var(--fs-md);margin-bottom:var(--sp-3)">SKILLS</h3>
<%          Dim gi, sj, gName
            For gi = 0 To rsSkillGrpCnt - 1
                gName = AV(rsSkillGrp, 0, gi) %>
            <div class="info-list" style="margin-bottom:var(--sp-2)">
                <div class="row">
                    <span class="lbl" style="flex:0 0 96px"><%= HE(gName) %></span>
                    <span>
<%                  For sj = 0 To rsSkillCnt - 1
                        If AV(rsSkill, SK_GROUP, sj) = gName Then %><span class="chip"><%= AVH(rsSkill, SK_NAME, sj) %></span><% End If
                    Next %>
                    </span>
                </div>
            </div>
<%          Next %>
        </section>
<% End If %>

<%'--- 학력 ---%>
<% If rsEduCnt > 0 Then %>
        <section class="mb-5">
            <h3 style="font-size:var(--fs-md);margin-bottom:var(--sp-3)">EDUCATION</h3>
            <ul class="timeline">
<%          Dim ei
            For ei = 0 To rsEduCnt - 1 %>
                <li>
                    <span class="period"><%= HE(DateUtil.Period(AV(rsEdu, CR_START, ei), AV(rsEdu, CR_END, ei), AVB(rsEdu, CR_CURRENT, ei))) %></span>
                    <h4><%= AVH(rsEdu, CR_TITLE, ei) %></h4>
                    <span class="position"><%= AVH(rsEdu, CR_POSITION, ei) %></span>
                </li>
<%          Next %>
            </ul>
        </section>
<% End If %>

<%'--- 자격 / 수상 ---%>
<% If rsCertCnt > 0 Or rsAwardCnt > 0 Then %>
        <section>
            <h3 style="font-size:var(--fs-md);margin-bottom:var(--sp-3)">CERTIFICATES &amp; AWARDS</h3>
            <ul class="timeline">
<%          Dim ti
            For ti = 0 To rsCertCnt - 1 %>
                <li>
                    <span class="period"><%= HE(DateUtil.Fmt(AV(rsCert, CR_START, ti), "ym")) %></span>
                    <h4><%= AVH(rsCert, CR_TITLE, ti) %></h4>
                    <span class="position"><%= AVH(rsCert, CR_POSITION, ti) %></span>
                </li>
<%          Next
            Dim ai
            For ai = 0 To rsAwardCnt - 1 %>
                <li>
                    <span class="period"><%= HE(DateUtil.Fmt(AV(rsAward, CR_START, ai), "ym")) %></span>
                    <h4><%= AVH(rsAward, CR_TITLE, ai) %></h4>
                    <span class="position"><%= AVH(rsAward, CR_POSITION, ai) %></span>
                </li>
<%          Next %>
            </ul>
        </section>
<% End If %>

    </article>

</div>

<!--#include file="../include/Footer.asp"-->
