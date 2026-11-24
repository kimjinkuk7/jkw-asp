<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/about/Career.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap wrap-narrow">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li>ABOUT</li>
            <li>경력 · 학력</li>
        </ul>
        <h2>경력 · 학력</h2>
        <p class="desc">총 경력 <strong><%= HE(crYears) %>년</strong></p>
    </div>

    <!--#include file="../include/Lnb.asp"-->

<%'--- 경력 ---%>
    <section class="section">
        <div class="section-head"><h3>경력</h3></div>
<% If crWorkCnt = 0 Then %>
        <div class="empty">등록된 경력이 없습니다.</div>
<% Else %>
        <ul class="timeline">
<%      Dim wi
        For wi = 0 To crWorkCnt - 1 %>
            <li>
                <span class="period"><%= HE(DateUtil.Period(AV(crWork, CR_START, wi), AV(crWork, CR_END, wi), AVB(crWork, CR_CURRENT, wi))) %></span>
                <h4>
<%                  If Len(AV(crWork, CR_ORGURL, wi)) > 0 Then %>
                    <a href="<%= HE(AV(crWork, CR_ORGURL, wi)) %>" target="_blank" rel="noopener noreferrer"><%= AVH(crWork, CR_TITLE, wi) %> ↗</a>
<%                  Else %>
                    <%= AVH(crWork, CR_TITLE, wi) %>
<%                  End If %>
<%                  If AVB(crWork, CR_CURRENT, wi) Then %><span class="badge badge-primary is-current">재직중</span><% End If %>
                </h4>
                <span class="position"><%= AVH(crWork, CR_POSITION, wi) %></span>
<%              If Len(AV(crWork, CR_DESC, wi)) > 0 Then %>
                <div class="desc"><%= AVH(crWork, CR_DESC, wi) %></div>
<%              End If %>
            </li>
<%      Next %>
        </ul>
<% End If %>
    </section>

<%'--- 학력 ---%>
<% If crEduCnt > 0 Then %>
    <section class="section">
        <div class="section-head"><h3>학력</h3></div>
        <ul class="timeline">
<%      Dim ei
        For ei = 0 To crEduCnt - 1 %>
            <li>
                <span class="period"><%= HE(DateUtil.Period(AV(crEdu, CR_START, ei), AV(crEdu, CR_END, ei), AVB(crEdu, CR_CURRENT, ei))) %></span>
                <h4><%= AVH(crEdu, CR_TITLE, ei) %></h4>
                <span class="position"><%= AVH(crEdu, CR_POSITION, ei) %></span>
<%              If Len(AV(crEdu, CR_DESC, ei)) > 0 Then %>
                <div class="desc"><%= AVH(crEdu, CR_DESC, ei) %></div>
<%              End If %>
            </li>
<%      Next %>
        </ul>
    </section>
<% End If %>

<%'--- 자격증 ---%>
<% If crCertCnt > 0 Then %>
    <section class="section">
        <div class="section-head"><h3>자격증</h3></div>
        <ul class="timeline">
<%      Dim ti
        For ti = 0 To crCertCnt - 1 %>
            <li>
                <span class="period"><%= HE(DateUtil.Fmt(AV(crCert, CR_START, ti), "ym")) %></span>
                <h4><%= AVH(crCert, CR_TITLE, ti) %></h4>
                <span class="position"><%= AVH(crCert, CR_POSITION, ti) %></span>
            </li>
<%      Next %>
        </ul>
    </section>
<% End If %>

<%'--- 수상 ---%>
<% If crAwardCnt > 0 Then %>
    <section class="section">
        <div class="section-head"><h3>수상</h3></div>
        <ul class="timeline">
<%      Dim ai
        For ai = 0 To crAwardCnt - 1 %>
            <li>
                <span class="period"><%= HE(DateUtil.Fmt(AV(crAward, CR_START, ai), "ym")) %></span>
                <h4><%= AVH(crAward, CR_TITLE, ai) %></h4>
                <span class="position"><%= AVH(crAward, CR_POSITION, ai) %></span>
<%              If Len(AV(crAward, CR_DESC, ai)) > 0 Then %>
                <div class="desc"><%= AVH(crAward, CR_DESC, ai) %></div>
<%              End If %>
            </li>
<%      Next %>
        </ul>
    </section>
<% End If %>

<%'--- 기타 활동 ---%>
<% If crActCnt > 0 Then %>
    <section class="section">
        <div class="section-head"><h3>활동</h3></div>
        <ul class="timeline">
<%      Dim yi
        For yi = 0 To crActCnt - 1 %>
            <li>
                <span class="period"><%= HE(DateUtil.Period(AV(crAct, CR_START, yi), AV(crAct, CR_END, yi), AVB(crAct, CR_CURRENT, yi))) %></span>
                <h4><%= AVH(crAct, CR_TITLE, yi) %></h4>
                <span class="position"><%= AVH(crAct, CR_POSITION, yi) %></span>
<%              If Len(AV(crAct, CR_DESC, yi)) > 0 Then %>
                <div class="desc"><%= AVH(crAct, CR_DESC, yi) %></div>
<%              End If %>
            </li>
<%      Next %>
        </ul>
    </section>
<% End If %>

    <div class="btn-row center mt-5">
        <a class="btn btn-primary" href="<%= U("/about/Resume.asp") %>">이력서로 보기</a>
        <a class="btn" href="<%= U("/about/Skill.asp") %>">기술 스택</a>
    </div>

</div>

<!--#include file="../include/Footer.asp"-->
