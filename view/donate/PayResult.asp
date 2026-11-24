<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/donate/PayResult.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="auth-wrap">

<% If prOk Then %>
        <h2>후원해 주셔서 감사합니다 ☕</h2>
        <p class="desc">받은 커피값은 밤샘 연료로 성실히 소비하겠습니다.</p>
<% Else %>
        <h2>결제가 완료되지 않았습니다</h2>
        <p class="desc">마음만으로도 충분히 감사합니다.</p>
<% End If %>

        <div class="auth-box">
            <div class="pay-result <%= IIfStr(prOk, "is-ok", "is-fail") %>">
                <div class="mark"><%= IIfStr(prOk, "✓", "✕") %></div>
                <div class="txt">
                    <%= IIfStr(prOk, "결제 완료", "결제 실패") %>
<% If Not prOk And Len(prMsg) > 0 Then %>
                    <div class="sub"><%= HE(prMsg) %></div>
<% End If %>
                </div>
            </div>

            <dl class="pay-summary mt-3">
                <dt>주문번호</dt><dd class="mono"><%= HE(prOrder) %></dd>
                <dt>후원자</dt>  <dd><%= HE(prName) %></dd>
                <dt>금액</dt>    <dd class="amt"><%= StringUtil.Comma(prAmount) %>원</dd>
<% If prOk Then %>
                <dt>결제수단</dt><dd><%= HE(prMethod) %></dd>
                <dt>거래번호</dt><dd class="mono fs-xs"><%= HE(prTid) %></dd>
<% End If %>
<% If AVB(prArr, DN_ISTEST, 0) Then %>
                <dt>구분</dt>    <dd><span class="badge">테스트 결제 · 실제 청구 없음</span></dd>
<% End If %>
            </dl>

            <div class="btn-row mt-4">
                <a href="<%= U("/donate/Default.asp") %>" class="btn btn-block">지원금 페이지로</a>
                <a href="<%= U("/main/Default.asp") %>" class="btn btn-block mt-2">메인으로</a>
            </div>
        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
