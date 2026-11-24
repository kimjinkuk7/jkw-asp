<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Login.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="auth-wrap">
        <h2>관리자 로그인</h2>
        <p class="desc"><%= HE(GLB_SITE_TITLE) %> 관리 시스템</p>

        <div class="auth-box" style="background:var(--c-bg)">
<% If Len(adErr) > 0 Then %>
            <div class="alert alert-error"><%= HE(adErr) %></div>
<% End If %>

            <form method="post" action="<%= U("/admin/Login.asp") %>" data-validate>
                <%= Auth.TokenField() %>
                <input type="hidden" name="rurl" value="<%= HE(adRurl) %>">

                <div class="form-group">
                    <label class="form-label" for="a-id">아이디</label>
                    <input class="input" type="text" id="a-id" name="login_id" maxlength="30"
                           autocomplete="username" autofocus data-required data-label="아이디">
                </div>
                <div class="form-group">
                    <label class="form-label" for="a-pw">비밀번호</label>
                    <input class="input" type="password" id="a-pw" name="user_pw" maxlength="64"
                           autocomplete="current-password" data-required data-label="비밀번호">
                </div>

                <button type="submit" class="btn btn-primary btn-block btn-lg">로그인</button>
            </form>

            <div class="auth-links">
                <a href="<%= U("/main/Default.asp") %>">사이트로 돌아가기</a>
            </div>
        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
