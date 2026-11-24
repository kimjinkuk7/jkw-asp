<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Stat.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>접속 통계</h2>
                <p class="desc">최근 <%= stDays %>일 기준 · PV = 페이지뷰, UV = 순 방문자(IP)</p>
            </div>

            <nav class="cate-tabs">
                <a href="<%= U("/admin/Stat.asp") %>?days=7"  class="<%= IIfStr(stDays = 7,  "is-on", "") %>">7일</a>
                <a href="<%= U("/admin/Stat.asp") %>?days=30" class="<%= IIfStr(stDays = 30, "is-on", "") %>">30일</a>
                <a href="<%= U("/admin/Stat.asp") %>?days=90" class="<%= IIfStr(stDays = 90, "is-on", "") %>">90일</a>
            </nav>

            <div class="dash-grid">
                <div class="dash-card"><div class="lbl">총 PV</div><div class="num"><%= StringUtil.Comma(stTotalPv) %></div></div>
                <div class="dash-card"><div class="lbl">총 UV</div><div class="num"><%= StringUtil.Comma(stTotalUv) %></div></div>
                <div class="dash-card"><div class="lbl">일평균 PV</div><div class="num"><%= StringUtil.Comma(stAvg) %></div></div>
                <div class="dash-card"><div class="lbl">집계 일수</div><div class="num"><%= stCnt %></div></div>
            </div>

            <div class="admin-card">
                <h3>기기 비율</h3>
<% If stDevCnt = 0 Then %>
                <div class="empty">데이터가 없습니다.</div>
<% Else %>
                <div style="display:flex;height:26px;border-radius:var(--radius-full);overflow:hidden;border:1px solid var(--c-border)">
<%      Dim di2, dpct, dcolors
        dcolors = Array("#2b5cff", "#10b981", "#f59e0b", "#e5484d")
        For di2 = 0 To stDevCnt - 1
            dpct = Int(AVI(stDev, 1, di2) * 100 / stDevTotal) %>
                    <div style="width:<%= dpct %>%;background:<%= dcolors(di2 Mod 4) %>" title="<%= AVH(stDev, 0, di2) %> <%= dpct %>%"></div>
<%      Next %>
                </div>
                <div class="mt-3">
<%      For di2 = 0 To stDevCnt - 1 %>
                    <span class="chip" style="border-left:4px solid <%= dcolors(di2 Mod 4) %>">
                        <%= AVH(stDev, 0, di2) %> <%= StringUtil.Comma(AVI(stDev, 1, di2)) %>
                        (<%= Int(AVI(stDev, 1, di2) * 100 / stDevTotal) %>%)
                    </span>
<%      Next %>
                </div>
<% End If %>
            </div>

            <div class="admin-card">
                <h3>일별 추이</h3>
<% If stCnt = 0 Then %>
                <div class="empty">방문 기록이 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead><tr><th>날짜</th><th class="num">PV</th><th class="num">UV</th><th style="width:50%">그래프</th></tr></thead>
                        <tbody>
<%      Dim ri
        For ri = 0 To stCnt - 1 %>
                            <tr>
                                <td class="fw7"><%= AVH(stArr, 0, ri) %></td>
                                <td class="num"><%= StringUtil.Comma(AVI(stArr, 1, ri)) %></td>
                                <td class="num"><%= StringUtil.Comma(AVI(stArr, 2, ri)) %></td>
                                <td>
                                    <div class="skill-bar" style="height:12px">
                                        <i style="width:<%= Int(AVI(stArr, 1, ri) * 100 / stMax) %>%"></i>
                                    </div>
                                </td>
                            </tr>
<%      Next %>
                        </tbody>
                    </table>
                </div>
<% End If %>
            </div>

            <div class="admin-card">
                <h3>집계 관리</h3>
                <p class="fs-sm sub">
                    <code>tb_visit_log</code> 는 원본 로그, <code>tb_visit_stat</code> 는 일별 집계 테이블입니다.<br>
                    운영 시에는 SQL Server 에이전트로 매일 새벽에
                    <code>EXEC dbo.usp_visit_stat_aggregate</code> 와
                    <code>EXEC dbo.usp_visit_log_purge 90</code> 을 실행하도록 등록하세요.
                </p>
                <div class="btn-row mt-3">
                    <a class="btn" href="<%= U("/admin/Stat.asp") %>?days=<%= stDays %>&amp;agg=1&amp;_csrf=<%= UE(Auth.Token()) %>"
                       data-confirm="오늘자 통계를 지금 집계할까요?">지금 집계 실행</a>
                </div>
            </div>

        </div>
    </div>
</div>

<%
'--- 수동 집계 요청 처리 (조회 화면 하단에서 실행) ---
If RequestUtil.ValueInit("agg", "") = "1" Then
    Auth.CheckToken RequestUtil.ValueInit("_csrf", "")
    Visit.Aggregate_p
    Response.Write "<script>alert('집계를 실행했습니다.');location.replace('" & JE(U("/admin/Stat.asp") & "?days=" & stDays) & "');</script>"
End If
%>
<!--#include file="../include/Footer.asp"-->
