<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/board/PwCheck.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="auth-wrap">
        <h2>비밀번호 확인</h2>
        <p class="desc">비회원이 작성한 게시물입니다. <%= HE(pcLabel) %>하려면 작성 시 입력한 비밀번호가 필요합니다.</p>

        <div class="auth-box">
            <p class="fs-sm mb-4"><strong><%= HE(StringUtil.Cut(pcTitle, 40)) %></strong></p>

<% If Len(pcErr) > 0 Then %>
            <div class="alert alert-error"><%= HE(pcErr) %></div>
<% End If %>

            <form method="post" action="<%= U("/board/PwCheck.asp") %>" data-validate>
                <%= Auth.TokenField() %>
                <input type="hidden" name="bid" value="<%= HE(bdId) %>">
                <input type="hidden" name="id"  value="<%= pcId %>">
                <input type="hidden" name="act" value="<%= HE(pcAct) %>">

                <div class="form-group">
                    <label class="form-label" for="pw">비밀번호</label>
                    <input class="input" type="password" id="pw" name="pw" maxlength="20"
                           autocomplete="current-password" autofocus
                           data-required data-label="비밀번호">
                </div>

                <button type="submit" class="btn btn-primary btn-block btn-lg">확인</button>
            </form>

            <div class="auth-links">
                <a href="<%= U("/board/List.asp") %>?bid=<%= UE(bdId) %>">목록으로</a>
            </div>
        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
